// PwC Tax Sharing Agreement Generator
// Template ID 3082 - Division 721 ITAA 1997 Compliant
// Last Updated: 2025

document.getElementById('tsaForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    const loadingDiv = document.getElementById('loading');
    loadingDiv.classList.add('active');
    
    try {
        const formData = collectFormData();
        const doc = await generateTSADocument(formData);
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().split('T')[0];
        const filename = `Tax_Sharing_Agreement_${formData.headCompany.name.replace(/[^a-z0-9]/gi, '_')}_${timestamp}.docx`;
        
        // Save the document
        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, filename);
        
        loadingDiv.classList.remove('active');
        alert('Tax Sharing Agreement generated successfully!');
        
    } catch (error) {
        console.error('Error generating document:', error);
        loadingDiv.classList.remove('active');
        alert('Error generating document. Please check the console for details.');
    }
});

function collectFormData() {
    // Collect all form data
    const data = {
        consolidationDate: document.getElementById('consolidationDate').value,
        noticeDate: document.getElementById('noticeDate').value,
        financialYearEnd: document.getElementById('financialYearEnd').value,
        isMEC: document.querySelector('input[name="isMEC"]:checked').value === 'yes',
        governingLaw: document.getElementById('governingLaw').value,
        
        headCompany: {
            name: document.getElementById('headCompanyName').value,
            abn: document.getElementById('headCompanyABN').value,
            acn: document.getElementById('headCompanyACN').value,
            address: document.getElementById('headCompanyAddress').value,
            email: document.getElementById('headCompanyEmail').value,
            attention: document.getElementById('headCompanyAttention').value
        },
        
        allocationMethod: document.querySelector('input[name="allocationMethod"]:checked').value,
        hasPriorTSA: document.querySelector('input[name="hasPriorTSA"]:checked').value === 'yes',
        priorTSADate: document.getElementById('priorTSADate')?.value || '',
        hasDOCG: document.querySelector('input[name="hasDOCG"]:checked').value === 'yes',
        includeDPT: document.querySelector('input[name="includeDPT"]')?.checked || false,
        
        members: collectMembers()
    };
    
    return data;
}

function collectMembers() {
    const members = [];
    const memberNames = document.getElementsByName('memberName[]');
    
    for (let i = 0; i < memberNames.length; i++) {
        members.push({
            name: memberNames[i].value,
            abn: document.getElementsByName('memberABN[]')[i].value,
            acn: document.getElementsByName('memberACN[]')[i].value,
            type: document.getElementsByName('memberType[]')[i].value,
            trustName: document.getElementsByName('trustName[]')[i]?.value || '',
            trustABN: document.getElementsByName('trustABN[]')[i]?.value || '',
            address: document.getElementsByName('memberAddress[]')[i].value,
            email: document.getElementsByName('memberEmail[]')[i].value
        });
    }
    
    return members;
}

async function generateTSADocument(data) {
    const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, UnderlineType, 
            NumberFormat, Table, TableRow, TableCell, WidthType, BorderStyle, Header } = docx;

    // Create document sections
    const sections = [{
        properties: {},
        children: [
            // Title
            new Paragraph({
                text: "Tax Sharing Agreement",
                heading: HeadingLevel.TITLE,
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            }),
            
            // Parties
            new Paragraph({
                children: [
                    new TextRun({
                        text: data.headCompany.name,
                        bold: true
                    }),
                    new TextRun({
                        text: ` (ABN ${data.headCompany.abn})`,
                    })
                ],
                spacing: { after: 200 }
            }),
            
            new Paragraph({
                text: "The Parties Listed in Schedule 1",
                bold: true,
                spacing: { after: 400 }
            }),
            
            // Table of Contents heading
            new Paragraph({
                text: "Contents",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 400, after: 200 }
            }),
            
            // Generate TOC entries
            ...generateTableOfContents(),
            
            // Date and Parties section
            new Paragraph({
                text: "Date",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 }
            }),
            
            new Paragraph({
                text: `________________, being the date of final execution`,
                spacing: { after: 200 }
            }),
            
            new Paragraph({
                text: "Parties",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 }
            }),
            
            ...generatePartiesSection(data),
            
            // Recitals
            new Paragraph({
                text: "Recitals",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 }
            }),
            
            ...generateRecitals(data),
            
            // PART A - INTERPRETATION
            new Paragraph({
                text: "PART A - INTERPRETATION",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 600, after: 300 }
            }),
            
            // Clause 1: Interpretation
            new Paragraph({
                text: "1. Interpretation",
                heading: HeadingLevel.HEADING_2,
                numbering: {
                    reference: "main-numbering",
                    level: 0
                },
                spacing: { before: 400, after: 200 }
            }),
            
            new Paragraph({
                text: "1.1 Definitions",
                heading: HeadingLevel.HEADING_3,
                numbering: {
                    reference: "main-numbering",
                    level: 1
                },
                spacing: { before: 200, after: 200 }
            }),
            
            new Paragraph({
                text: "Unless otherwise defined in this Agreement, capitalised terms used in this Agreement which have a defined meaning in the ITAA have the same meaning in this Agreement and:",
                spacing: { after: 200 }
            }),
            
            ...generateDefinitions(data),
            
            new Paragraph({
                text: "1.2 Construction",
                heading: HeadingLevel.HEADING_3,
                numbering: {
                    reference: "main-numbering",
                    level: 1
                },
                spacing: { before: 400, after: 200 }
            }),
            
            ...generateConstructionRules(),
            
            // PART B - ALLOCATION
            new Paragraph({
                text: "PART B - ALLOCATION",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 600, after: 300 }
            }),
            
            // Clause 2: Allocation of Group Liability
            ...generateClause2(data),
            
            // Clause 3: Calculation and Recording
            ...generateClause3(),
            
            // Clause 4: Clear Exit
            ...generateClause4(data),
            
            // PART C - MISCELLANEOUS
            new Paragraph({
                text: "PART C - MISCELLANEOUS",
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 600, after: 300 }
            }),
            
            // Clause 5: Provision of TSA to Commissioner
            ...generateClause5(),
            
            // Clause 6: Tax Audits and Disputes
            ...generateClause6(),
            
            // Clause 7: Dispute Resolution
            ...generateClause7(),
            
            // Clause 8: Changes to the Law
            ...generateClause8(),
            
            // Clause 9: Precedence of Regulations
            ...generateClause9(),
            
            // Clause 10: Prior TSA (if applicable)
            ...(data.hasPriorTSA ? generateClause10(data) : []),
            
            // Clause 11: Miscellaneous
            ...generateClause11(data),
            
            // Clause 12: Notices
            ...generateClause12(),
            
            // Clause 13: Representations and Warranties
            ...generateClause13(data),
            
            // Schedule 1: Contributing Members
            ...generateSchedule1(data),
            
            // Schedule 2: Allocation Principles
            ...generateSchedule2(data),
            
            // Schedule 3: Accession Agreement
            ...generateSchedule3(data),
            
            // Schedule 4: Release Agreement
            ...generateSchedule4(data),
            
            // Execution Pages
            ...generateExecutionPages(data)
        ]
    }];

    return new Document({
        sections: sections,
        numbering: {
            config: [{
                reference: "main-numbering",
                levels: [
                    {
                        level: 0,
                        format: NumberFormat.DECIMAL,
                        text: "%1.",
                        alignment: AlignmentType.LEFT
                    },
                    {
                        level: 1,
                        format: NumberFormat.DECIMAL,
                        text: "%1.%2",
                        alignment: AlignmentType.LEFT
                    },
                    {
                        level: 2,
                        format: NumberFormat.LOWER_LETTER,
                        text: "(%3)",
                        alignment: AlignmentType.LEFT
                    }
                ]
            }]
        }
    });
}

function generateTableOfContents() {
    const { Paragraph, TextRun } = docx;
    
    const contents = [
        "1. Interpretation",
        "2. Allocation of Group Liability",
        "3. Calculation and Recording of Contribution Amounts",
        "4. Clear Exit From Group Liability",
        "5. Provision of Tax Sharing Agreement to Commissioner",
        "6. Tax Audits and Disputes",
        "7. Dispute Resolution",
        "8. Changes to the Law",
        "9. Precedence of Regulations",
        "10. Prior Tax Sharing Agreement",
        "11. Miscellaneous",
        "12. Notices",
        "13. Representations and Warranties",
        "Schedule 1 -- Contributing Members",
        "Schedule 2 -- Allocation Principles",
        "Schedule 3 -- Accession Agreement",
        "Schedule 4 -- Release Agreement",
        "Execution Pages"
    ];
    
    return contents.map(item => new Paragraph({
        text: item,
        spacing: { after: 100 }
    }));
}

function generatePartiesSection(data) {
    const { Paragraph, Table, TableRow, TableCell, WidthType } = docx;
    
    return [
        new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph("Name")],
                            width: { size: 30, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: data.headCompany.name, bold: true })]
                        })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("ABN")] }),
                        new TableCell({ children: [new Paragraph({ text: data.headCompany.abn, bold: true })] })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Description")] }),
                        new TableCell({ children: [new Paragraph({ text: "Head Company", bold: true })] })
                    ]
                }),
                new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph("Notice details")] }),
                        new TableCell({ 
                            children: [
                                new Paragraph(data.headCompany.address),
                                new Paragraph(`E-mail: ${data.headCompany.email}`),
                                new Paragraph(`Attention: ${data.headCompany.attention}`)
                            ] 
                        })
                    ]
                })
            ]
        })
    ];
}

function generateRecitals(data) {
    const { Paragraph, TextRun } = docx;
    
    const recitalA = data.isMEC 
        ? `The Contributing Members and the Head Company are members of a multiple entry consolidated group for Australian tax purposes formed originally with effect from ${formatDate(data.consolidationDate)}.`
        : `The Contributing Members and the Head Company are members of a consolidated group for Australian tax purposes formed originally with effect from ${formatDate(data.consolidationDate)}.`;
    
    return [
        new Paragraph({
            children: [
                new TextRun({ text: "A. ", bold: true }),
                new TextRun(recitalA)
            ],
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "B. ", bold: true }),
                new TextRun("Under section 721-15 of ITAA 1997, the Head Company and each of the Contributing Members are jointly and severally liable to pay a Group Liability if the Head Company were to fail to meet that Group Liability by the Due Date, unless the Group Liability is covered by a Tax Sharing Agreement.")
            ],
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "C. ", bold: true }),
                new TextRun("While acknowledging that the primary obligation to pay any Group Liability rests with the Head Company, the Group Members wish to enter into this Agreement to ensure that the Group Members are not jointly and severally liable for a Group Liability if the Head Company fails to pay such Group Liability by the relevant Due Date, and intend this Agreement to be a Tax Sharing Agreement.")
            ],
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "D. ", bold: true }),
                new TextRun("The Group Members have not entered into this Agreement as part of an arrangement with the purpose of prejudicing the recovery by the Commissioner of some or all of a Group Liability.")
            ],
            spacing: { after: 400 }
        })
    ];
}

function generateDefinitions(data) {
    const { Paragraph, TextRun } = docx;
    
    const consolidatedGroupDef = data.isMEC
        ? "Consolidated Group means the MEC of which the Head Company is the provisional head company appointed in accordance with subsection 719-60 of ITAA 1997."
        : "Consolidated Group means the consolidated group as defined in subsection 703-5 of ITAA 1997 of which the Head Company is the head company for the purposes of subsection 703-15 of ITAA 1997.";
    
    const definitions = [
        { term: "Accession Agreement", def: "means an accession agreement substantially in the form of Schedule 3 or in such other form determined by the Head Company." },
        { term: "Agreement", def: "means this Tax Sharing Agreement and all schedules, annexures and attachments to it, as amended by the parties in accordance with its terms." },
        { term: "Allocation Principles", def: "means the allocation principles set out in Schedule 2." },
        { term: "Amended Assessment", def: "means an assessment which is amended by the Commissioner in accordance with section 170 of ITAA 1936." },
        { term: "Business Day", def: `means a day other than a Saturday, Sunday or public holiday on which banks are open for general business in ${data.governingLaw}.` },
        { term: "Calculation Advice", def: "means a notice provided by the Head Company to a Group Member under clause 3.1." },
        { term: "Capital Gain", def: "has the meaning given to that term in the ITAA 1997." },
        { term: "CGT Event", def: "has the meaning given to that term in the ITAA 1997." },
        { term: "Commissioner", def: "means the Federal Commissioner of Taxation of the Commonwealth of Australia or the Australian Taxation Office." },
        { term: "Consolidated Group", def: consolidatedGroupDef },
        { term: "Contributing Member", def: "means, subject to the terms of this Agreement, an entity, Trust or Partnership listed in Schedule 1 of this Agreement and any other entity, Trust or Partnership which becomes a party to this Agreement by executing an Accession Agreement." },
        { term: "Contribution Amount", def: "means the amount determined in accordance with Part B of this Agreement for each Group Member in relation to a Group Liability." },
        { term: "Corporations Act", def: "means the Corporations Act 2001 (Cth) as amended and in force from time to time." },
        { term: "Dispute Expert", def: "has the meaning given to the term in clause 7.3." },
        { term: "Due Date", def: "in relation to a Group Liability, means the time at which the Group Liability becomes, or became, due and payable by the Head Company to the Commissioner." },
        { term: "Electronic Signature", def: "means: (a) an encrypted signature applied using a proprietary program (for example DocuSign or AdobeSign) which is applied following verification of an individual's identity; or (b) the digital image of an individual's manuscript signature (regardless of whether it is a digitally generated image, or a scanned copy of a physically signed document)." },
        { term: "Exit Date", def: "has the meaning given to the term in clause 4.2." },
        { term: "Exiting Member", def: "has the meaning given to the term in clause 4.1." },
        { term: "Financial Year", def: `means a period of 12 months ending on ${formatDate(data.financialYearEnd)} each year provided that the first Financial Year of the Consolidated Group will commence on the date specified in Recital A.` },
        { term: "Government Agency", def: "means any government or governmental, semi‑governmental, administrative, fiscal or judicial body, department, commission, authority, tribunal, agency or entity whether foreign, federal, state, territorial or local." },
        { term: "Group Liability", def: "means any Tax-Related Liability of the Head Company of the Consolidated Group including any Amended Assessment of such Tax-Related Liability where a basis for determining an allocation of that Tax-Related Liability is specified in Schedule 2." },
        { term: "Group Member", def: "means each of the Head Company and a Contributing Member." },
        { term: "GST Act", def: "means A New Tax System (Goods and Services Tax) Act 1999 (Cth), as in force from time to time." },
        { term: "Head Company", def: `means ${data.headCompany.name} (ABN ${data.headCompany.abn})${data.isMEC ? ' or such other Group Member as becomes the provisional head company of the Consolidated Group in accordance with clause 11.24.' : ''}` },
        { term: "ITAA", def: "means, where applicable, ITAA 1936 or ITAA 1997." },
        { term: "ITAA 1936", def: "means the Income Tax Assessment Act 1936 (Cth)." },
        { term: "ITAA 1997", def: "means the Income Tax Assessment Act 1997 (Cth)." },
        { term: "Item", def: "means an item or items specified in column 1 of the table in subsection 721-10(2) of ITAA 1997." },
        { term: "Item 3 Tax-Related Liability", def: "means the Tax-Related Liability as referred to in item 3 of the table in subsection 721-10(2) of the ITAA 1997." },
        { term: "Notice", def: "has the meaning given to that term in clause 12(a)." },
        { term: "Partnership", def: "means a partnership as defined in subsection 995-1 of ITAA 1997." },
        { term: "Release Agreement", def: "means a release agreement substantially in the form of Schedule 4 or in such other form determined by the Head Company." },
        { term: "Tax Sharing Agreement or TSA", def: "means a tax sharing agreement for the purposes of Division 721 of ITAA 1997, as amended from time to time." },
        { term: "Tax-Related Liability", def: "of the Head Company has the same meaning as it has in subsection 721-10(2) of ITAA 1997." }
    ];
    
    if (data.includeDPT) {
        definitions.push(
            { term: "DPT Base Amount", def: "has the meaning set out in section 177P(2) of the ITAA 1936." },
            { term: "DPT Group Liability", def: "means the Group Liability referred to in Item 115 of section 721-10(2) of the ITAA 1997." },
            { term: "DPT Tax Benefit", def: "means a tax benefit as defined in section 177J(1) of the ITAA 1936." },
            { term: "Significant Global Entity", def: "has the same meaning as it has in section 960-555 of the ITAA 1997." }
        );
    }
    
    return definitions.map(def => new Paragraph({
        children: [
            new TextRun({ text: `${def.term} `, bold: true }),
            new TextRun(def.def)
        ],
        spacing: { after: 100 }
    }));
}

function generateConstructionRules() {
    const { Paragraph } = docx;
    
    const rules = [
        "words importing the singular include the plural and vice versa;",
        "words importing a gender include any gender;",
        "where a word or phrase is given a particular meaning, other parts of speech and grammatical forms of a word or phrase defined in this Agreement have a corresponding meaning;",
        "an expression importing a natural person includes any individual, company, partnership, trust, joint venture, association, corporation or other body corporate and any Government Agency;",
        "no provision of this Agreement will be construed adversely to a party solely on the ground that the party was responsible for the preparation of this Agreement or that provision;",
        "when the day on which something must be done is not a Business Day, that thing must be done on the preceding Business Day;",
        "a reference to a monetary amount is a reference to Australian Dollars."
    ];
    
    return [
        new Paragraph({
            text: "In this Agreement headings are for convenience only and do not affect the interpretation of this Agreement and, unless the context otherwise requires:",
            spacing: { after: 200 }
        }),
        ...rules.map((rule, index) => new Paragraph({
            text: `${String.fromCharCode(97 + index)}. ${rule}`,
            spacing: { after: 100 }
        }))
    ];
}

function generateClause2(data) {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "2. Allocation of Group Liability",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "2.1 Consideration",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "In consideration for the mutual promises set out in this Agreement, each Group Member agrees to be bound by the terms and conditions set out in this Agreement.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.2 Purpose of Allocation Provisions",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The purpose of the allocation provisions of this Agreement, as contained in this Part B (and supplemented by the other provisions of this Agreement), is to:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. reasonably allocate each Group Liability among the Group Members;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. provide certainty in relation to such allocation so that a Group Member's Contribution Amount may be determined primarily by reference to this Part B; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. meet any and all requirements for a valid Tax Sharing Agreement.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.3 Reasonable Allocation",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The amount of the Group Liability allocated to each Group Member under this Agreement represents a reasonable allocation of the total amount of the Group Liability among the Group Members immediately prior to the Due Date for that Group Liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.4 Allocation of Group Liabilities",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Subject to clause 2.5, each Group Liability with an original assessment with a Due Date after the date of this Agreement must be allocated either to, or among, the Head Company and each of the Contributing Members who were parties to this Agreement immediately prior to the Due Date of the original assessment of such Group Liability in accordance with the relevant Allocation Principles.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.5 Pre-Consolidation Liabilities",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Any liabilities of a Group Member to the Commissioner that relate to the period prior to that party becoming a Member of the Consolidated Group remain the liability of that Group Member.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.6 Payment to Commissioner",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Subject to the terms of this Agreement, if the Head Company does not pay or otherwise discharge in full a Group Liability by the Due Date, each Contributing Member will only be liable to pay the Commissioner an amount equal to the Contribution Amount of that Contributing Member for that Group Liability, and such amount will be due and payable to the Commissioner, in the circumstances set out in section 721-30(5) or (5A) of ITAA 1997, as the case may be.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.7 No Debt",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "No Group Member may claim a debt is due from another Group Member in relation to any Contribution Amount payable by a Contributing Member to the Commissioner.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "2.8 No Reduction for Funding Amount",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The amount payable by a Contributing Member in accordance with clause 2.6 will not be reduced, increased or otherwise affected by any amount paid, payable, received or receivable by a Contributing Member under the terms of a Tax Funding Agreement or any other tax funding or contribution agreement between one or more of the Group Members.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause3() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "3. Calculation and Recording of Contribution Amounts",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "3.1 Calculation Advice",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Immediately upon the earlier of the Head Company becoming aware of a possible default, or actual default, in the payment of part or all of a Group Liability by the relevant Due Date, the Head Company must calculate the Contribution Amount of each Group Member in respect of that Group Liability and advise each Group Member in writing of:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. the amount of its Contribution Amount;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. the method of calculation of the amount in paragraph (a) above; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. any other information it considers reasonable.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "The Contribution Amount of each Group Member as calculated by the Head Company under this clause 3.1 is binding on each Group Member except in the case of manifest error.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "3.2 Assistance by Contributing Members",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Each Contributing Member must provide the Head Company with any information it may reasonably require in relation to a Group Liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "3.3 Retention of Records",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Each Group Member must retain a record of each Calculation Advice (provided to, or by, it, as the case may be) for a period of not less than seven years from the date of its making.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "3.4 Provision of Calculation Advices",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Subject to the terms of this Agreement, each Contributing Member is only entitled to be given Calculation Advices that relate to that particular Contributing Member and not any Calculation Advices that relate to any other Group Member.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "3.5 Annexures",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: 'The Head Company must attach as Annexure "A" to its counterpart of this Agreement any:',
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. Calculation Advice issued by it to a Group Member under clause 3.1;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. Accession Agreement entered into by it with a new Group Member;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. Release Agreement entered into by it with an Exiting Member; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "d. other document it considers relevant to the interpretation of this Agreement.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause4(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    const baseContent = [
        new Paragraph({
            text: "4. Clear Exit From Group Liability",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "4.1 Exit from the Consolidated Group",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "This clause 4 applies if a Group Member will cease to be a Member of the Consolidated Group (Exiting Member).",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.2 Provision of information to Exiting Member",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Prior to the date the Exiting Member ceases to be a Member of the Consolidated Group (Exit Date), the Head Company must provide the Exiting Member with a calculation (which for the avoidance of doubt may be nil) of:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. if a Contribution Amount can be determined for the Exiting Member pursuant to clause 2 for a Group Liability of the Head Company with a Due Date after the Exit Date, that Contribution Amount;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. if a Contribution Amount cannot be determined for the Exiting Member pursuant to clause 2 for a Group Liability of the Head Company with a Due Date after the Exit Date, a reasonable estimate of that Contribution Amount;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. if, before the Exit Date, the Head Company anticipates that after the Exit Date there will be an Amended Assessment in respect of a Tax-Related Liability, which relates to a Tax-Related Period during which the Exiting Member was a Member of the Consolidated Group, an amount that is a reasonable estimate of the Contribution Amount of the Exiting Member that takes into account the anticipated Amended Assessment; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "d. such other amount that Division 721 of ITAA 1997 may specify from time to time that will allow the Exiting Member to leave the Consolidated Group clear of a Group Liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.3 Conclusive Evidence",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The calculation by the Head Company of any Contribution Amount due under clause 4.2 is conclusive evidence of that Contribution Amount of the Exiting Member in respect to a Group Liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.4 Payment of Contribution Amount by Exiting Member",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Before the Exit Date, the Exiting Member must pay the Head Company the Contribution Amounts, if any, calculated under clause 4.2, at which time the Head Company must acknowledge receipt of such payment in writing and hold such payment separate from the other funds of the Head Company until the Due Date for each relevant Group Liability at which time the payment may be used to meet the relevant Tax-Related Liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.5 Nil Contribution Amount",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "For the avoidance of doubt, if a Contribution Amount calculated pursuant to clause 4.2 is nil, the Exiting Member will be deemed to have made the relevant payment under clause 4.4.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.6 Release",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Head Company and the Exiting Member must execute a Release Agreement:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. provided that the Exiting Member has made; or",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. that is effective conditional upon the Exiting Member making,",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "all payments required under clause 4.4. Each Group Member appoints the Head Company as its agent to enter into the Release Agreement referred to in this clause 4.6.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.7 Acknowledgment",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Despite anything else in this Agreement but without affecting any party's rights or obligations under this Agreement, pursuant to ITAA 1997 an Exiting Member may be liable to pay the Commissioner part or all of a Group Liability in certain circumstances as outlined in Division 721 of the ITAA 1997.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.8 Provision of Agreement to Exiting Member",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Despite anything else in this Agreement, an Exiting Member may at any time, after the Exit Date, request that the Head Company provide it with a copy of this Agreement and the Head Company must immediately on such demand provide that Exiting Member with a copy of this Agreement in the approved form and any Annexures to it together with any other documents or information necessary for its validity as a TSA under Division 721 of ITAA 1997.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "4.9 No Prejudice",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Head Company and the Exiting Member agree and acknowledge that the Exiting Member does not, or will not, cease to be a Group Member under an arrangement, a purpose of which is to prejudice the recovery by the Commissioner of some or all of the amount of any Group Liability.",
            spacing: { after: 300 }
        })
    ];
    
    // Add MEC-specific clause if applicable
    if (data.isMEC) {
        baseContent.push(
            new Paragraph({
                text: "4.10 Exit by Provisional Head Company",
                heading: HeadingLevel.HEADING_3,
                spacing: { before: 200, after: 100 }
            }),
            
            new Paragraph({
                text: "Subject to clause 4.10(b), where the Exiting Member is or was a provisional head company of the Consolidated Group in a relevant Financial Year, a reference to Head Company in this clause 4 is to be read as a reference to the Group Member that is the provisional head company of the Consolidated Group at the time that Exiting Member leaves the Consolidated Group.",
                spacing: { after: 200 }
            })
        );
    }
    
    return baseContent;
}

function generateClause5() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "5. Provision of Tax Sharing Agreement to Commissioner",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "5.1 Obligation to Provide",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Head Company must provide the Commissioner with a copy of this Agreement in the approved form and any Annexures to it together with any other documents or information necessary for its validity as a TSA under Division 721 of ITAA 1997 within 14 days of notice given by the Commissioner pursuant to subsection 721-25(3) of ITAA 1997 or as otherwise required.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "5.2 Extension of Time",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If the Head Company is unable to provide to the Commissioner a copy of this Agreement in accordance with clause 5.1, the Head Company must apply, pursuant to ITAA 1997, to the Commissioner for an extension of time to lodge a copy of this Agreement and other documents contemplated in clause 5.1.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "5.3 Indemnity",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Head Company indemnifies each Member of the Consolidated Group for any Group Liability or other amount incurred or increased as a result of its failure to comply with its obligations under clause 5.1 or 5.2.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "5.4 Provision of Notice to other Group Members",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Within three days of the receipt of a notice given by the Commissioner pursuant to subsection 721-25(3) of ITAA 1997, the Head Company must give a copy of such notice to all other Group Members.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "5.5 Provision by Contributing Member as Agent of Head Company",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Despite anything in this clause 5, if the Head Company is given a notice by the Commissioner pursuant to subsection 721-25(3) of ITAA 1997 and a Contributing Member provides a copy of this Agreement, any Annexures to it or documents or information referred to in clause 5.1 to the Commissioner within the time required by that notice, the Contributing Member is, by this Agreement, appointed agent of the Head Company for the purpose of complying with that notice.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "5.6 Provision by Contributing Member",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If a Contributing Member is given a notice by the Commissioner pursuant to subsection 721-15(5) of ITAA 1997, the Head Company must immediately on demand provide that Contributing Member with a copy of this Agreement in the approved form and any Annexures to it together with any other documents or information necessary for its validity as a TSA under Division 721 of ITAA 1997.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause6() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "6. Tax Audits and Disputes",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "6.1 Audit",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If the Commissioner commences an examination of, inquiry into or audit of a Group Liability (Tax Audit) for any Tax-Related Period in respect of which a Contributing Member was part of the Consolidated Group, that Contributing Member must provide such documents, information, explanations and assistance to the Head Company relating to the Group Liability or the Tax Audit as the Head Company may reasonably request and must fully co-operate with the Head Company in its management of the Tax Audit.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "6.2 Disputes",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If any dispute (Tax Dispute), including but not limited to any legal proceedings, arises in relation to a Group Liability for any Tax-Related Period in respect of which a Contributing Member was part of the Consolidated Group, that Contributing Member must provide such documents, information, explanations and assistance to the Head Company relating to the Group Liability or the Tax Dispute as the Head Company requires in its management of the Tax Dispute.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "6.3 Provision of Information to Contributing Member",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If any Contributing Member becomes liable to pay to the Commissioner the whole or any portion of a Group Liability for any Tax-Related Period, either by reason of section 721-15 of ITAA 1997 or under this Agreement, the Head Company must provide such information to the Contributing Member relating to the Group Liability as the Contributing Member may reasonably request and must fully co-operate with the Contributing Member in relation to payment of, or disputing, the liability.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "6.4 Agency",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Each Contributing Member appoints the Head Company as its agent in respect of all matters concerning its compliance with the ITAA or related legislation (including, but not limited to, Tax Audits and Tax Disputes) and covenants with the Head Company and with each other Contributing Member that it will not knowingly, and otherwise than in the ordinary course of its business, make any statement or do, or omit to do, any thing (including disclosing to the Commissioner any information or document) (Tax-Related Disclosure) that could reasonably be expected to have the effect of increasing any Group Liability of the Head Company or any Contribution Amount of another Contributing Member, without the express prior approval of the Head Company.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "6.5 Disclosure by Law",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If a Contributing Member is required by any law, or by notice given under any law, to make a Tax-Related Disclosure to another person, clause 6.4 does not apply to the making of that Tax-Related Disclosure, provided that the Contributing Member has used its best endeavours to obtain the prior approval of the Head Company to make that Tax-Related Disclosure.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "6.6 Communications with Commissioner",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Head Company may make any disclosures, requests for amendment of any return or similar communication to the Commissioner as it considers appropriate, having regard to its legal obligations.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause7() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "7. Dispute Resolution",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "7.1 Dispute Resolution",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Any dispute between the parties in relation to this Agreement, or the rights and/or obligations arising pursuant to it, to the extent it is not governed specifically by another provision of this Agreement, must be dealt with in accordance with this clause 7.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "7.2 Notice of Dispute",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Any party who claims that a dispute within the meaning of clause 7.1 has arisen must, within five Business Days of becoming aware of such dispute, serve written notice upon all other parties to this Agreement in accordance with clause 12 and in such notice provide full particulars of the dispute and any/or all claims made by that party and alleged to arise pursuant to the dispute.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "7.3 Referral of Dispute",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Upon receipt of a notice pursuant to clause 7.2, the Head Company must refer the dispute to any person agreed to by the parties to the dispute or, failing such agreement, an expert nominated and appointed in accordance with, and subject to, the Resolution Institute Expert Determination Rules (the Dispute Expert).",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "7.4 Expeditious Determination",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The parties must use best endeavours to procure that the Dispute Expert expeditiously determines the subject matter of the dispute in such manner as he or she reasonably determines.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "7.5 Binding Decision",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Except in the case of manifest error, the parties are bound by the decision of the Dispute Expert, including any order as to costs.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "7.6 Court Proceedings",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "This clause 7 does not constrain the Head Company from commencing proceedings in a Court of competent jurisdiction against any Group Member who fails to make a payment required by clause 4 for recovery of such payment as a liquidated debt.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause8() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "8. Changes to the Law",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "8.1 Parties to Act in Good Faith",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If a change occurs (including, but not limited to for the purpose of this clause 8, a change that occurs because a regulation is made for the purposes of section 721-25(1)(d) of ITAA 1997) to a law or policy of the Commissioner that applies to this Agreement in a manner that affects the legal or practical effect that, but for that change, this Agreement would have had, or results in this Agreement not covering all Group Liabilities for the purposes of section 721-25 of ITAA 1997, then as soon as practicable after becoming aware of the change, the Contributing Members must act in good faith with the Head Company to determine and make such changes (if any) to this Agreement as may be necessary.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "8.2 Deemed Amendment",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Despite clause 8.1, if a change to a law or policy of the Commissioner occurs that applies to this Agreement in a manner that results in this Agreement not covering all Group Liabilities for the purposes of section 721-25 of ITAA 1997, this Agreement is deemed to be amended from the date such law came into effect, so that the relevant Group Liability not covered will be covered and allocated among the Group Members on the same basis as if it were an Item 3 Tax-Related Liability.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause9() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "9. Precedence of Regulations",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "9.1 Variations",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The parties must make such variations to this Agreement as may be necessary to ensure this Agreement complies with any regulations made under paragraph 721-25(1)(d) of ITAA 1997 or is a valid Tax Sharing Agreement within the contemplation of section 721-25 of ITAA 1997 and any policy of the Commissioner.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "9.2 No Agreement",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If the parties are unable to agree on any variations as contemplated by clause 9.1, the parties must appoint a Dispute Expert. The Dispute Expert may decide the necessary variations and the Dispute Expert's decision is final and binding on the parties (except in the case of manifest error).",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "9.3 Variation by Dispute Expert",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "This Agreement is deemed to be varied in accordance with the decision of the Dispute Expert effective from the date specified in the decision. The Group Members must share the costs of the Dispute Expert equally, unless prior agreement can be reached between the parties.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "9.4 Deemed Incorporation of rules and regulations",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Subject to clause 9.4(b), if the Commissioner or the Treasury issues any ruling or makes any regulation stipulating any matter which must or must not be included in a TSA, or which states what will be acceptable as a means of calculating a liability to income tax, or which in any other way regulates what a TSA must contain or must not contain, then this Agreement is deemed to incorporate any such ruling or regulation to the extent necessary to preserve or achieve validity of this Agreement as a TSA under Division 721 of ITAA 1997.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "If the ruling or regulation prescribes alternatives for preserving or achieving such validity, this Agreement is deemed to incorporate the alternative most consistent with the pre-existing terms of this Agreement.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause10(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "10. Prior Tax Sharing Agreement",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "10.1 Relationship between TSAs",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: `The Tax Sharing Agreement entered into between the Head Company and other members of the Consolidated Group on ${formatDate(data.priorTSADate)} (Prior TSA):`,
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. does not apply to any Group Liability for which the Head Company's original Due Date is after the date of this Agreement; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. despite anything in this Agreement, continues to apply to a Group Liability with a Due Date after the date of this Agreement where the original Due Date of that Group Liability was on or before the date of this Agreement.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "This Agreement:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. applies to each Group Liability with an original Due Date after the date of this Agreement; and",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. despite anything in this Agreement, does not apply to any Group Liability in relation to which the Prior TSA applies.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause11(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "11. Miscellaneous",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "11.1 Amendment",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "The Group Members (other than any Exiting Member released pursuant to clause 4.6 or a Group Member that has been liquidated, deregistered, wound up or dissolved) may at any time amend, vary or replace this Agreement by written agreement, but not so as to adversely affect the obligations or rights of any Exiting Member released pursuant to clause 4.6.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.2 GST",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Unless otherwise expressly stated, the consideration to be provided or payment obligation under this Agreement is exclusive of GST. If GST is imposed on any supply made under this Agreement, the recipient must pay to the supplier any amount equal to the GST payable on the supply.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.3 Additional Contributing Members",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "If a new Member joins the Consolidated Group, the Head Company must procure the new Member to promptly execute an Accession Agreement and perform any actions necessary to ensure that the new Member's membership of the Consolidated Group does not prejudice the validity of this Agreement as a TSA.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.4 Termination",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "This Agreement commences on the date of this Agreement and continues until terminated by mutual agreement of the parties or the Due Date for all Group Liabilities has passed after the Consolidated Group ceases to exist. The termination of this Agreement does not affect the rights, obligations and liabilities of any party that have accrued up to the date of termination.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.5 Governing Law and Jurisdiction",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: `This Agreement is governed by the laws of ${data.governingLaw}. Each party irrevocably and unconditionally submits to the non-exclusive jurisdiction of the courts of ${data.governingLaw}.`,
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.6 Severability",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Any provision of, or the application of any provision of this Agreement, which is prohibited, void, illegal or unenforceable in any jurisdiction is, in that jurisdiction, ineffective only to the extent to which it is void, illegal, unenforceable or prohibited.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.7 Entire Agreement",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: `This Agreement embodies the entire agreement between the parties with respect to the subject matter of this Agreement${data.hasPriorTSA ? ', other than as provided for in clause 10,' : ''} and supersedes any prior negotiation, arrangement, understanding or agreement with respect to the subject matter or any term of this Agreement.`,
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.8 Counterparts",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "This Agreement may be executed in any number of counterparts. All counterparts, taken together, constitute one instrument. A party may execute this Agreement by signing any counterpart.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "11.9 Electronic Signature",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "A party (including any signatory for a party) may execute this Agreement by Electronic Signature. Each party consents to the use of Electronic Signature and acknowledges that the use of Electronic Signature is an appropriately reliable method for the purposes of this Agreement.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause12() {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "12. Notices",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "12.1 Form of communication",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Unless expressly stated otherwise in this Agreement, any notice, certificate, consent, request, demand, approval, waiver or other communication (Notice) must be in legible writing and in English, signed by the sender, and marked for the attention of and addressed to the addressee.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "12.2 Delivery of Notices",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Notices must be hand delivered or sent by prepaid express post (next day delivery) or email to the addressee's address for notices specified in the 'Parties' section or Schedule 1 of this Agreement or to any other address or email a party notifies to the other parties under this clause 12.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "12.3 When Notice is effective",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Notices take effect from the time they are received or taken to be received under clause 12.4 (whichever happens first) unless a later time is specified.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "12.4 When Notice taken to be received",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Notice is taken to be received by the addressee:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. if by delivery in person, when delivered to the addressee;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. if by prepaid express post, on the second Business Day after the date of posting;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. if by electronic mail (e-mail), four hours after the sent time (as recorded on the sender's e-mail server), unless the sender receives a notice that the message has not been delivered.",
            spacing: { after: 300 }
        })
    ];
}

function generateClause13(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "13. Representations and Warranties",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 }
        }),
        
        new Paragraph({
            text: "13.1 General Representations",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "Each party represents and warrants to each other party:",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. it is validly existing under the laws of the place of its incorporation or creation;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. it has the power to enter into and perform its obligations under this Agreement;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. it has taken all necessary action to authorise the entry into and performance of this Agreement;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "d. this Agreement is valid and binding and enforceable against it in accordance with its terms;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "e. the execution and performance by it of this Agreement does not and will not violate in any respect a provision of any law or its constituent documents;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "f. it has not entered into this Agreement as part of an arrangement any purpose of which was to prejudice the recovery by the Commissioner of some or all of the amount of any Group Liability;",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "g. it is aware that its allocation of a Group Liability under this Agreement may be more, or less, than the taxation amounts it would otherwise be liable for if it were not a Member of the Consolidated Group.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "13.2 Head Company Warranty",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: `The Head Company warrants that it satisfies the statutory requirements specified for a head company of a Consolidated Group as set out in section ${data.isMEC ? '719-65' : '703-15'} of ITAA 1997.`,
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "13.3 Contributing Member Warranty",
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: `Each Contributing Member warrants that it satisfies the statutory requirements specified for a subsidiary member for a Consolidated Group as set out in section ${data.isMEC ? '719-10 or section 719-5' : '703-15'} of ITAA 1997.`,
            spacing: { after: 300 }
        })
    ];
}

function generateSchedule1(data) {
    const { Paragraph, HeadingLevel, Table, TableRow, TableCell, WidthType } = docx;
    
    const memberRows = data.members.map(member => {
        const entityDisplay = member.type === 'trust' 
            ? `${member.name} (ACN ${member.acn}) as trustee for ${member.trustName} (ABN ${member.trustABN})`
            : member.type === 'partnership'
            ? `${member.name} (ACN ${member.acn}) as general partner for ${member.trustName}`
            : `${member.name} (ACN ${member.acn})`;
            
        return new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph(entityDisplay)],
                    width: { size: 40, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [new Paragraph(member.abn || 'N/A')],
                    width: { size: 20, type: WidthType.PERCENTAGE }
                }),
                new TableCell({
                    children: [
                        new Paragraph(member.address),
                        new Paragraph(`Email: ${member.email}`)
                    ],
                    width: { size: 40, type: WidthType.PERCENTAGE }
                })
            ]
        });
    });
    
    return [
        new Paragraph({
            text: "SCHEDULE 1 -- CONTRIBUTING MEMBERS",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 600, after: 300 },
            pageBreakBefore: true
        }),
        
        new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ text: "Member Name", bold: true })],
                            width: { size: 40, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: "ABN / ACN", bold: true })],
                            width: { size: 20, type: WidthType.PERCENTAGE }
                        }),
                        new TableCell({
                            children: [new Paragraph({ text: "Address and E-mail", bold: true })],
                            width: { size: 40, type: WidthType.PERCENTAGE }
                        })
                    ]
                }),
                ...memberRows
            ]
        })
    ];
}

function generateSchedule2(data) {
    const { Paragraph, HeadingLevel, TextRun } = docx;
    
    const baseAllocations = [
        new Paragraph({
            text: "SCHEDULE 2 -- ALLOCATION PRINCIPLES",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 600, after: 300 },
            pageBreakBefore: true
        }),
        
        new Paragraph({
            text: "The Allocation Principles for each of the Tax-Related Liabilities of the Head Company are as set out below.",
            spacing: { after: 300 }
        }),
        
        // Item 5: Untainting Tax
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Item 5 subsection 721-10(2) of ITAA 1997 -- section 197-70 of ITAA 1997 (untainting tax).")
            ],
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun("The Group Member which made the election under section 197-55 of ITAA 1997 must pay the whole of this Tax-Related Liability.")
            ],
            spacing: { after: 300 }
        }),
        
        // Items 10, 15, 20, 22: Franking Tax
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Items 10, 15, 20, and 22 subsection 721-10(2) of ITAA 1997 (franking tax).")
            ],
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun("The Head Company must pay the entire amount of each of these Tax-Related Liabilities.")
            ],
            spacing: { after: 300 }
        }),
        
        // Item 3: Income Tax
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Item 3 subsection 721-10(2) of ITAA 1997 -- section 5-5 of ITAA 1997 (income tax and other amounts treated in the same way as income tax under that section).")
            ],
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun(`The Contribution Amount for each Group Member in respect to the Item 3 Tax-Related Liability will be determined based on a ${data.allocationMethod === 'notional_tax' ? 'Notional Taxation' : 'Notional Accounting'} Basis as set out below:`)
            ],
            spacing: { after: 200 }
        })
    ];
    
    // Add the appropriate allocation methodology
    if (data.allocationMethod === 'notional_tax') {
        baseAllocations.push(...generateNotionalTaxAllocation());
    } else {
        baseAllocations.push(...generateNotionalAccountingAllocation());
    }
    
    // Add PAYG Instalments allocation
    baseAllocations.push(
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Items 30, 32, and 35 subsection 721-10(2) of ITAA 1997 (PAYG instalments).")
            ],
            spacing: { before: 300, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun("The Tax-Related Liability will be allocated between each Group Member in proportion to their Income Tax Instalment Incomes for the relevant Tax-Related Period as determined in accordance with section 45-120 of Schedule 1 of TAA 1953.")
            ],
            spacing: { after: 300 }
        })
    );
    
    // Add GIC and Penalties allocation
    baseAllocations.push(
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Items 40, 45, 50, 55, 60, 65, and 70 subsection 721-10(2) of ITAA 1997 (general interest charge, administrative penalties, and shortfall interest charge).")
            ],
            spacing: { before: 300, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun("The Tax-Related Liability will be allocated between each Group Member based upon the allocation that was adopted for the underlying Item to which the relevant interest charge or penalty relates.")
            ],
            spacing: { after: 300 }
        })
    );
    
    // Add DPT allocation if applicable
    if (data.includeDPT) {
        baseAllocations.push(...generateDPTAllocation());
    }
    
    return baseAllocations;
}

function generateNotionalTaxAllocation() {
    const { Paragraph, TextRun } = docx;
    
    return [
        new Paragraph({
            children: [
                new TextRun({ text: "Notional Taxation Basis", bold: true, underline: { type: UnderlineType.SINGLE } })
            ],
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "a. An amount reflecting the notional income, or loss, as the case may be, for taxation for each Member of the Consolidated Group for the relevant year of income will be determined assuming that each was a stand alone entity and not a Member of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. The notional income, or loss, of a Member who was not a party to this Agreement prior to the Due Date of the original assessment of the Tax-Related Liability will be added to the notional income, or loss, of the Head Company.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. If the determination in (a) and (b) above results in a Group Member having a notional loss then such loss will be notionally transferred to each Group Member with a notional income pro-rata in proportion to the amount by which such Group Member's notional income bears to the notional income of all Group Members whose notional income is greater than zero, and that Group Member's notional taxation liability will be zero.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "d. The notional taxation liability of a Member will be determined by multiplying the notional income of that Member by the tax rate applicable to the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "e. The Tax-Related Liability will be allocated among each Group Member in proportion to the amount which that Group Member's notional taxation liability bears to the aggregate notional taxation liability of all Group Members.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "f. The amount so allocated to each Group Member is its Gross Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "g. The Gross Income Tax Contribution Amount of each Group Member will then be reduced by any Contribution Amounts for Items 30, 32 and 35 Tax-Related Liabilities allocated to it in respect of that year of income and any other amounts paid or payable by that Group Member for which the Head Company has, does or will receive a credit under section 45‑30 in Schedule 1 to TAA 1953 for that income year or a credit under section 45‑865 in Schedule 1 to TAA 1953 for that income year.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "h. The amount so allocated to each Group Member is its Net Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "i. If any Group Member's Net Income Tax Contribution Amount is less than or equal to zero, such Net Income Tax Contribution Amount will be notionally transferred to each Group Member with a Net Income Tax Contribution Amount of greater than zero, pro rata, and that Group Member's Net Adjusted Income Tax Contribution Amount will be zero.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "j. The amount so allocated to each Group Member is its Net Adjusted Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "k. The Tax-Related Liability will then be reduced by the amount of any credit under section 45-30 in Schedule 1 of TAA 1953 for that income year and any credit under section 45-865 in Schedule 1 of TAA 1953 for that income year (the 721-25(1A) Adjusted Tax Related Liability).",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "l. The 721-25(1A) Adjusted Tax Related Liability will be allocated between each Group Member in proportion to the amount which that Group Member's Net Adjusted Income Tax Contribution Amount bears to the aggregate Net Adjusted Income Tax Contribution Amounts of all Group Members.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "m. The amount so allocated to each Group Member is its Contribution Amount.",
            spacing: { after: 300 }
        })
    ];
}

function generateNotionalAccountingAllocation() {
    const { Paragraph, TextRun } = docx;
    
    return [
        new Paragraph({
            children: [
                new TextRun({ text: "Notional Accounting Basis", bold: true, underline: { type: UnderlineType.SINGLE } })
            ],
            spacing: { before: 200, after: 100 }
        }),
        
        new Paragraph({
            text: "a. The notional profit or loss before tax of each Member of the Consolidated Group will be determined based on the accounts of each such Member in the relevant year of income.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. An adjustment will be made to the amount calculated at paragraph (a) to eliminate any transactions between Members and exclude, where the Member is a Trust or Partnership, any amounts that would be included in determining the profit before tax of another Member.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. Any income or expenses that are referrable to part of the relevant year of income in which the Member was not part of the Consolidated Group will be apportioned on a reasonable basis to the period during which the Member was not a Member of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "d. The notional profit or loss before tax of a Member who was not a party to this Agreement prior to the Due Date of the original assessment of the Tax-Related Liability will be added to notional profit or loss before tax of the Head Company.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "e. If after applying the above paragraphs a Group Member has a notional loss then such loss will be notionally transferred to each Group Member with a notional profit pro-rata, and that Group Member's profit or loss before tax will be zero.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "f. The Tax Related Liability will be allocated among each Group Member based on the proportion that the notional profit or loss before tax of each Group Member bears to the sum of the notional profit or loss before tax for all Group Members.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "g. The amount so allocated to each Group Member is its Gross Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "h. The Gross Income Tax Contribution Amount of each Group Member will then be reduced by any Contribution Amounts for Items 30, 32 and 35 Tax Related Liabilities allocated to it in respect of that year of income and any other amounts paid or payable by that Group Member for which the Head Company has, does or will receive a credit under section 45‑30 in Schedule 1 to TAA 1953 for that income year or a credit under section 45‑865 in Schedule 1 to TAA 1953 for that income year.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "i. The amount so allocated to each Group Member is its Net Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "j. If any Group Member's Net Income Tax Contribution Amount is less than or equal to zero, such Net Income Tax Contribution Amount will be notionally transferred to each Group Member with a Net Income Tax Contribution Amount of greater than zero, pro rata, and that Group Member's Net Adjusted Income Tax Contribution Amount will be zero.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "k. The amount so allocated to each Group Member is its Net Adjusted Income Tax Contribution Amount.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "l. The Tax-Related Liability will then be reduced by the amount of any credit under section 45-30 in Schedule 1 of TAA 1953 for that income year and any credit under section 45-865 in Schedule 1 of TAA 1953 for that income year (the 721-25(1A) Adjusted Tax-Related Liability).",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "m. The 721-25(1A) Adjusted Tax-Related Liability will be allocated between each Group Member in proportion to the amount which that Group Member's Net Adjusted Income Tax Contribution Amount bears to the aggregate Net Adjusted Income Tax Contribution Amounts of all Group Members.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "n. The amount so allocated to each Group Member is its Contribution Amount.",
            spacing: { after: 300 }
        })
    ];
}

function generateDPTAllocation() {
    const { Paragraph, TextRun } = docx;
    
    return [
        new Paragraph({
            children: [
                new TextRun({ text: "Tax-Related Liability: ", bold: true }),
                new TextRun("Item 115 of section 721-10(2) of the ITAA 1997 (diverted profits tax).")
            ],
            spacing: { before: 300, after: 100 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Allocation Principles: ", bold: true }),
                new TextRun("The Contribution Amount for each Group Member in respect to the Item 115 Tax-Related Liability is determined on the following basis:")
            ],
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "a. In determining the Contribution Amount, it will be assumed that each Member is a Significant Global Entity, section 701-1 of the ITAA 1997 did not apply to any Member of the Consolidated Group, and the Contribution Amount of the Head Company will incorporate the relevant attributes of a Member who was not a party to this Agreement prior to the Due Date.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "b. If the DPT Group Liability arose due to only one DPT Tax Benefit, then the Contribution Amount for a Group Member will be determined in accordance with the following formula: Contribution Amount = A/B × DPT Group Liability, where A is the DPT Base Amount for that DPT Tax Benefit that would have been obtained by that Group Member, and B is the sum of all DPT Base Amounts for all DPT Tax Benefits that all Members would have obtained.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "c. If the DPT Group Liability arose due to more than one DPT Tax Benefit, then the Contribution Amount for a Group Member will be determined by apportioning the DPT Group Liability among the DPT Tax Benefits pro rata to the respective DPT Base Amounts and working out each Group Member's Contribution Amount in relation to each such portion.",
            spacing: { after: 300 }
        })
    ];
}

function generateSchedule3(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "SCHEDULE 3 -- ACCESSION AGREEMENT",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 600, after: 300 },
            pageBreakBefore: true
        }),
        
        new Paragraph({
            text: "This Accession Agreement template allows new members to join the Tax Sharing Agreement after the initial execution date.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "Date: ________________",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Parties:", bold: true })
            ],
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "1. [Full name of acceding party] (Acceding Party)",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `2. ${data.headCompany.name} (Head Company)`,
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "3. Each Contributing Member as defined in the TSA",
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Recitals:", bold: true })
            ],
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `A. On [date], the Head Company entered into a Tax Sharing Agreement (TSA) with persons that were at that date Members of the Head Company's Consolidated Group.`,
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "B. The Acceding Party has since that date become a Member of the Head Company's Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "C. The Head Company (on its own behalf and as agent for each of the other parties to the TSA) and the Acceding Party have agreed that the Acceding Party should become a party to the TSA on the terms set out in this agreement.",
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Agreed Terms:", bold: true })
            ],
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "1. The Acceding Party confirms that it has been supplied with a copy of the TSA and agrees and covenants with all present parties to the TSA to observe, perform and be bound by all the terms of the TSA so that the Acceding Party is deemed from the date of this agreement to be a Contributing Member as defined in, and for the purposes of, the TSA.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "2. The Head Company and the Acceding Party undertake to perform their respective duties and obligations under the TSA in respect to Group Liabilities with a Due Date after the date of this agreement.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "3. The Acceding Party represents and warrants that it is validly existing, has the power to enter into this agreement, this agreement is valid and binding, and the execution does not violate any law or its constituent documents.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `4. This agreement is governed by the laws of ${data.governingLaw}.`,
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            text: "Executed as an agreement.",
            spacing: { after: 200 }
        })
    ];
}

function generateSchedule4(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    return [
        new Paragraph({
            text: "SCHEDULE 4 -- RELEASE AGREEMENT",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 600, after: 300 },
            pageBreakBefore: true
        }),
        
        new Paragraph({
            text: "This Release Agreement template is executed when a member exits the Consolidated Group.",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "Date: ________________",
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Parties:", bold: true })
            ],
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "1. [Full name of exiting party] (Exiting Party)",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `2. ${data.headCompany.name} (Head Company)`,
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "3. Each Contributing Member as defined in the TSA",
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Recitals:", bold: true })
            ],
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `A. On [date], the Head Company entered into a Tax Sharing Agreement (TSA) with persons that were at that date Members of the Head Company's Consolidated Group.`,
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "B. The Exiting Party will cease to be a Member of the Head Company's Consolidated Group on [date].",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "C. The Head Company (on its own behalf and as agent for each other party to the TSA) and the Exiting Party have agreed that the Exiting Party and the Head Company (and each other party to the TSA) will be released from certain obligations under the TSA on the terms set out in this agreement.",
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            children: [
                new TextRun({ text: "Agreed Terms:", bold: true })
            ],
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "1. Subject to the terms of the TSA, the Head Company (on its own behalf and on behalf of other party to the TSA) by this agreement releases the Exiting Party from its obligations under the TSA that relate to any Tax-Related Period during which the Exiting Party is or was, for the whole of the Tax-Related Period, not part of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "2. Subject to the terms of the TSA, the Exiting Party by this agreement releases the Head Company and each other party to the TSA from their respective obligations under the TSA that relate to any Tax-Related Period during which the Exiting Party is or was, for the whole of the Tax-Related Period, not part of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "3. For the avoidance of doubt, this agreement does not release any party to the TSA from its obligations under the TSA that relate to any Tax-Related Period during which the Exiting Party was part of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "4. The Head Company confirms that the Exiting Party has paid its clear exit payment payable under clause 4 of the TSA prior to the date the Exiting Party ceased to be a Member of the Consolidated Group.",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: `5. This agreement is governed by the laws of ${data.governingLaw}.`,
            spacing: { after: 300 }
        }),
        
        new Paragraph({
            text: "Executed as an agreement.",
            spacing: { after: 200 }
        })
    ];
}

function generateExecutionPages(data) {
    const { Paragraph, HeadingLevel } = docx;
    
    const pages = [
        new Paragraph({
            text: "EXECUTION PAGES",
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 600, after: 300 },
            pageBreakBefore: true
        }),
        
        new Paragraph({
            text: "Executed as an agreement",
            spacing: { after: 400 }
        }),
        
        // Head Company execution block
        new Paragraph({
            text: `SIGNED by ${data.headCompany.name.toUpperCase()} by two Directors or a Director and Secretary in accordance with section 127 of the Corporations Act 2001 (Cth):`,
            spacing: { after: 200 }
        }),
        
        new Paragraph({
            text: "Signature of Director: _______________________     Signature of Director/Secretary: _______________________",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "Full Name of Signatory: _____________________     Full Name of Signatory: _____________________",
            spacing: { after: 100 }
        }),
        
        new Paragraph({
            text: "Date: _______________                                Date: _______________",
            spacing: { after: 400 }
        })
    ];
    
    // Add execution blocks for each contributing member
    data.members.forEach(member => {
        pages.push(
            new Paragraph({
                text: `SIGNED by ${member.name.toUpperCase()} by two Directors or a Director and Secretary in accordance with section 127 of the Corporations Act 2001 (Cth):`,
                spacing: { after: 200 }
            }),
            
            new Paragraph({
                text: "Signature of Director: _______________________     Signature of Director/Secretary: _______________________",
                spacing: { after: 100 }
            }),
            
            new Paragraph({
                text: "Full Name of Signatory: _____________________     Full Name of Signatory: _____________________",
                spacing: { after: 100 }
            }),
            
            new Paragraph({
                text: "Date: _______________                                Date: _______________",
                spacing: { after: 400 }
            })
        );
    });
    
    return pages;
}

function formatDate(dateString) {
    if (!dateString) return '[Date]';
    const date = new Date(dateString);
    const options = { day: 'numeric', month: 'long', year: 'numeric' };
    return date.toLocaleDateString('en-AU', options);
}
