// Questionnaire Parser for PwC TSA Generator
// Extracts data from completed Word questionnaires

class QuestionnaireParser {
    constructor(text) {
        this.text = text;
        this.data = {
            consolidationDate: '',
            noticeDate: '',
            financialYearEnd: '',
            isMEC: false,
            governingLaw: 'New South Wales',
            headCompany: {
                name: '',
                abn: '',
                acn: '',
                address: '',
                email: '',
                attention: ''
            },
            allocationMethod: 'taxation',
            hasPriorTSA: false,
            priorTSADate: '',
            hasDOCG: false,
            includeDPT: false,
            members: []
        };
    }

    parse() {
        this.extractDates();
        this.extractGroupType();
        this.extractHeadCompany();
        this.extractMembers();
        this.extractAllocationMethod();
        this.extractOtherDetails();
        return this.data;
    }

    extractDates() {
        // Extract consolidation date
        const consolidationPatterns = [
            /consolidation.*?(\d{1,2}\s+\w+\s+\d{4})/i,
            /consolidate.*?from.*?(\d{1,2}\s+\w+\s+\d{4})/i,
            /Notice of Consolidation.*?(\d{1,2}\s+\w+\s+\d{4})/i
        ];

        for (const pattern of consolidationPatterns) {
            const match = this.text.match(pattern);
            if (match) {
                this.data.consolidationDate = this.formatDate(match[1]);
                break;
            }
        }

        // Extract financial year end
        const fyEndMatch = this.text.match(/year.*?end.*?(\d{1,2}\s+\w+\s+\d{4})/i);
        if (fyEndMatch) {
            this.data.financialYearEnd = fyEndMatch[1];
        }
    }

    extractGroupType() {
        // Check if MEC group
        const mecMatch = this.text.match(/MEC.*?group.*?(yes|no)/i) || 
                        this.text.match(/Multiple Entry.*?(yes|no)/i);
        if (mecMatch) {
            this.data.isMEC = mecMatch[1].toLowerCase() === 'yes';
        }

        // Check for DOCG
        const docgMatch = this.text.match(/Deed of Cross Guarantee.*?(yes|no)/i) ||
                         this.text.match(/DOCG.*?(yes|no)/i);
        if (docgMatch) {
            this.data.hasDOCG = docgMatch[1].toLowerCase() === 'yes';
        }

        // Check for prior TSA
        const priorTSAMatch = this.text.match(/previously entered.*?tax sharing.*?(yes|no)/i);
        if (priorTSAMatch) {
            this.data.hasPriorTSA = priorTSAMatch[1].toLowerCase() === 'yes';
        }
    }

    extractHeadCompany() {
        // Extract head company name
        const namePatterns = [
            /Head Company.*?Name.*?[:：]\s*([^\n]+)/i,
            /head company.*?[:：]\s*([A-Z][^\n]+Pty Ltd)/i,
            /TopCo.*?Pty Ltd/i
        ];

        for (const pattern of namePatterns) {
            const match = this.text.match(pattern);
            if (match) {
                this.data.headCompany.name = this.cleanText(match[1] || match[0]);
                break;
            }
        }

        // Extract ABN
        const abnMatch = this.text.match(/ABN.*?[:：]\s*(\d{2}\s*\d{3}\s*\d{3}\s*\d{3})/i);
        if (abnMatch) {
            this.data.headCompany.abn = abnMatch[1].replace(/\s/g, ' ');
        }

        // Extract ACN
        const acnMatch = this.text.match(/ACN.*?[:：]\s*(\d{3}\s*\d{3}\s*\d{3})/i);
        if (acnMatch) {
            this.data.headCompany.acn = acnMatch[1].replace(/\s/g, ' ');
        }

        // Extract address
        const addressMatch = this.text.match(/address.*?[:：]\s*([^\n]+(?:\n[^\n]+){0,2})/i);
        if (addressMatch) {
            this.data.headCompany.address = this.cleanText(addressMatch[1]);
        }

        // Extract email
        const emailMatch = this.text.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
        if (emailMatch) {
            this.data.headCompany.email = emailMatch[0];
        }
    }

    extractMembers() {
        // Look for member entities in the document
        const memberPatterns = [
            // Pattern 1: Full entity details
            /([A-Z][^\n]*Pty Ltd)\s*\(ACN\s*(\d{3}\s*\d{3}\s*\d{3})\)/g,
            // Pattern 2: Trust entities
            /([\w\s]+)\s*as trustee for\s*([\w\s]+Trust)/gi
        ];

        const foundMembers = new Set();

        for (const pattern of memberPatterns) {
            let match;
            while ((match = pattern.exec(this.text)) !== null) {
                const memberName = this.cleanText(match[1]);
                
                // Skip head company
                if (memberName === this.data.headCompany.name) {
                    continue;
                }

                // Skip duplicates
                if (foundMembers.has(memberName)) {
                    continue;
                }

                foundMembers.add(memberName);

                const member = {
                    name: memberName,
                    abn: '',
                    acn: match[2] ? match[2].replace(/\s/g, ' ') : '',
                    type: 'company',
                    trustName: '',
                    trustABN: '',
                    address: this.data.headCompany.address, // Default to head company address
                    email: ''
                };

                // Check if it's a trust
                if (match[0].includes('trustee')) {
                    member.type = 'trust';
                    member.trustName = match[2] || '';
                }

                this.data.members.push(member);
            }
        }

        // This method can be extended to add custom entity recognition if needed
        this.parseKnownEntities();
    }

    parseKnownEntities() {
        // This method can be customized by users to add their own entity recognition
        // For now, it relies on the general pattern matching in extractMembers()
        
        // Users can add custom entity patterns here if needed for their specific use case
        // Example format:
        // const customEntities = [
        //     { name: 'Entity Name Pty Ltd', acn: 'XXX XXX XXX' }
        // ];
    }

    extractAllocationMethod() {
        // Check allocation method selection
        if (this.text.match(/notional\s+tax(ation)?/i) && !this.text.match(/notional\s+accounting/i)) {
            this.data.allocationMethod = 'taxation';
        } else if (this.text.match(/notional\s+accounting/i)) {
            this.data.allocationMethod = 'accounting';
        }

        // The default is already 'taxation' in constructor
    }

    extractOtherDetails() {
        // Check for DPT inclusion
        const dptMatch = this.text.match(/Diverted Profits Tax.*?(yes|no)/i) ||
                        this.text.match(/DPT.*?provisions.*?(yes|no)/i);
        if (dptMatch) {
            this.data.includeDPT = dptMatch[1].toLowerCase() === 'yes';
        }

        // Extract governing law if specified
        const govLawMatch = this.text.match(/governing\s+law.*?[:：]\s*([^\n]+)/i);
        if (govLawMatch) {
            this.data.governingLaw = this.cleanText(govLawMatch[1]);
        }
    }

    formatDate(dateString) {
        // Try to parse and format the date
        const date = new Date(dateString);
        if (!isNaN(date.getTime())) {
            return date.toISOString().split('T')[0];
        }
        return '';
    }

    cleanText(text) {
        return text
            .replace(/\s+/g, ' ')
            .replace(/[*_]/g, '')
            .replace(/^\s+|\s+$/g, '')
            .replace(/\[.*?\]/g, '');
    }
}

// Function to populate form with parsed data
function populateFormWithQuestionnaireData(parsedData) {
    console.log('Populating form with:', parsedData);

    // Section A: Group Information
    if (parsedData.consolidationDate) {
        document.getElementById('consolidationDate').value = parsedData.consolidationDate;
    }
    if (parsedData.noticeDate) {
        document.getElementById('noticeDate').value = parsedData.noticeDate;
    }
    if (parsedData.financialYearEnd) {
        document.getElementById('financialYearEnd').value = parsedData.financialYearEnd;
    }
    if (parsedData.governingLaw) {
        document.getElementById('governingLaw').value = parsedData.governingLaw;
    }

    // Radio buttons
    document.querySelector(`input[name="isMEC"][value="${parsedData.isMEC ? 'yes' : 'no'}"]`).checked = true;
    document.querySelector(`input[name="hasDOCG"][value="${parsedData.hasDOCG ? 'yes' : 'no'}"]`).checked = true;
    document.querySelector(`input[name="hasPriorTSA"][value="${parsedData.hasPriorTSA ? 'yes' : 'no'}"]`).checked = true;

    // Prior TSA date if applicable
    if (parsedData.hasPriorTSA && parsedData.priorTSADate) {
        document.getElementById('priorTSADateGroup').style.display = 'block';
        document.getElementById('priorTSADate').value = parsedData.priorTSADate;
    }

    // Section B: Head Company
    if (parsedData.headCompany.name) {
        document.getElementById('headCompanyName').value = parsedData.headCompany.name;
    }
    if (parsedData.headCompany.abn) {
        document.getElementById('headCompanyABN').value = parsedData.headCompany.abn;
    }
    if (parsedData.headCompany.acn) {
        document.getElementById('headCompanyACN').value = parsedData.headCompany.acn;
    }
    if (parsedData.headCompany.address) {
        document.getElementById('headCompanyAddress').value = parsedData.headCompany.address;
    }
    if (parsedData.headCompany.email) {
        document.getElementById('headCompanyEmail').value = parsedData.headCompany.email;
    }
    if (parsedData.headCompany.attention) {
        document.getElementById('headCompanyAttention').value = parsedData.headCompany.attention;
    }

    // Section C: Allocation Method
    document.querySelector(`input[name="allocationMethod"][value="${parsedData.allocationMethod}"]`).checked = true;
    if (parsedData.includeDPT) {
        document.getElementById('includeDPT').checked = true;
    }

    // Section D: Contributing Members
    const membersContainer = document.getElementById('membersContainer');
    membersContainer.innerHTML = ''; // Clear existing members
    memberCount = 0;

    parsedData.members.forEach((member, index) => {
        addMember(); // Add new member card
        
        const currentIndex = memberCount;
        
        // Populate member fields
        document.getElementsByName('memberName[]')[index].value = member.name;
        document.getElementsByName('memberABN[]')[index].value = member.abn;
        document.getElementsByName('memberACN[]')[index].value = member.acn;
        document.getElementsByName('memberType[]')[index].value = member.type;
        document.getElementsByName('memberAddress[]')[index].value = member.address;
        document.getElementsByName('memberEmail[]')[index].value = member.email;

        // If trust, populate trust fields
        if (member.type === 'trust') {
            const trustFields = document.getElementById(`trustFields-${currentIndex}`);
            if (trustFields) {
                trustFields.style.display = 'block';
                if (document.getElementsByName('trustName[]')[index]) {
                    document.getElementsByName('trustName[]')[index].value = member.trustName;
                }
                if (document.getElementsByName('trustABN[]')[index]) {
                    document.getElementsByName('trustABN[]')[index].value = member.trustABN;
                }
            }
        }
    });

    console.log('Form populated successfully');
}

// Export for use in main app
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { QuestionnaireParser, populateFormWithQuestionnaireData };
}
