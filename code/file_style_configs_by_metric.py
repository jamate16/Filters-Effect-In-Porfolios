from file_style import FileStyle, FileStyleDetails, FileStyleDetailsFactory

file_style_a_details_factory = FileStyleDetailsFactory("%b-%Y", "A1", " | ", "A1", "A3", "C6")
file_style_b_details_factory = FileStyleDetailsFactory("%d-%m-%Y", "B2", " (", "A1", "A14", "B15")

file_style_configs_by_metric = {
    "Pretax ROA": {
        FileStyle.A: file_style_a_details_factory.create("Pretax ROA", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Pretax ROA", "Financial Summary")
    },
    "Gross Margin": {
        FileStyle.A: file_style_a_details_factory.create("Gross Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Gross Profit Margin", "Financial Summary")
    },
    "EBITDA Margin": {
        FileStyle.A: file_style_a_details_factory.create("EBITDA Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("EBITDA Margin", "Financial Summary")
    },
    "Operating Margin": {
        FileStyle.A: file_style_a_details_factory.create("Operating Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Operating Margin", "Financial Summary")
    },
    "Net Margin": {
        FileStyle.A: file_style_a_details_factory.create("Net Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Net Margin", "Financial Summary")
    },
    "Net Margin": {
        FileStyle.A: file_style_a_details_factory.create("Net Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Net Margin", "Financial Summary")
    },
    "Pretax ROE": {
        FileStyle.A: file_style_a_details_factory.create("Pretax ROE", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Pretax ROE", "Financial Summary")
    },
    "ROIC": {
        FileStyle.A: file_style_a_details_factory.create("ROIC", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Pretax ROE", "Financial Summary")
    },
}
