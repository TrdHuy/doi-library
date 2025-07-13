from data.doi_template.v1.slide_data_model import DocumentContent, SectionContent, SlideContent, SlideBlock, SlideBlockItem
from data.doi_template.v1.contracts.md_contract import SectionId  # enum anh vừa tạo

MOCK_DOCUMENT = DocumentContent(
    sections=[
        SectionContent(
            section_id=SectionId.TECHNICAL_FIELD.value,
            slides=[
                SlideContent(blocks=[
                    SlideBlock(
                        heading="Overview of DGAs",
                        items=[
                            SlideBlockItem(text="Domain Generation Algorithms (DGAs) are used by botnets to generate domain names algorithmically.", type="paragraph"),
                            SlideBlockItem(text="These algorithms help malware evade traditional detection systems.", type="paragraph"),
                            SlideBlockItem(text="Random domain patterns", type="bullet", level=0),
                            SlideBlockItem(text="Hard to blacklist", type="bullet", level=0),
                            SlideBlockItem(text="Adaptive to takedown", type="bullet", level=0),
                        ],
                        image_path="../_assets/images/dga_overview.png"
                    )
                ])
            ]
        ),
        SectionContent(
            section_id=SectionId.PROBLEM.value,
            slides=[
                SlideContent(blocks=[
                    SlideBlock(
                        heading="Challenges in Detecting DGAs",
                        items=[
                            SlideBlockItem(text="Character-based DGAs generate domain names that lack semantic meaning.", type="paragraph"),
                            SlideBlockItem(text="This makes them difficult to distinguish from legitimate domains.", type="paragraph"),
                            SlideBlockItem(text="Non-human-readable domains", type="bullet", level=0),
                            SlideBlockItem(text="No lexical patterns", type="bullet", level=0),
                            SlideBlockItem(text="Frequent daily updates", type="bullet", level=0),
                        ]
                    )
                ])
            ]
        ),
        SectionContent(
            section_id=SectionId.PRIOR_ART.value,
            slides=[
                SlideContent(blocks=[
                    SlideBlock(
                        heading="Blacklist Approaches",
                        items=[
                            SlideBlockItem(text="Traditional blacklists fail to keep up with the dynamic nature of DGA-based domains.", type="paragraph"),
                            SlideBlockItem(text="Some systems use manual updates, which are slow and ineffective.", type="paragraph")
                        ]
                    ),
                    SlideBlock(
                        heading="Machine Learning Attempts",
                        items=[
                            SlideBlockItem(text="Early ML-based solutions used lexical features but lacked time series or frequency analysis.", type="paragraph"),
                            SlideBlockItem(text="Rely on static features", type="bullet", level=0),
                            SlideBlockItem(text="Poor generalization", type="bullet", level=0),
                        ],
                        image_path="../_assets/images/ml_attempts.png"
                    )
                ])
            ]
        ),
        SectionContent(
            section_id=SectionId.PURPOSE.value,
            slides=[
                SlideContent(blocks=[
                    SlideBlock(
                        heading="Our Proposed Method",
                        items=[
                            SlideBlockItem(text="We propose a two-stage DGA detection system leveraging snapshot features and frequency modeling.", type="paragraph"),
                            SlideBlockItem(text="This method improves detection accuracy and reduces false positives.", type="paragraph"),
                            SlideBlockItem(text="Combines static and dynamic features", type="bullet", level=0),
                            SlideBlockItem(text="Utilizes time-aware models", type="bullet", level=0),
                        ]
                    )
                ])
            ]
        )
    ]
)
