from ..data_fetchers.file_style import FileStyle, FileStyleDetails, FileStyleDetailsFactory

file_style_a_details_factory = FileStyleDetailsFactory("A1", " | ", "A1", "A3", "C5")
file_style_b_details_factory = FileStyleDetailsFactory("B2", " (", "A1", "A14", "B11")

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
    "Pretax Margin": {
        FileStyle.A: file_style_a_details_factory.create("Pretax Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Income Before Tax Margin", "Financial Summary")
    },
    "Net Margin": {
        FileStyle.A: file_style_a_details_factory.create("Net Margin", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Net Margin", "Financial Summary")
    },
    "Pretax ROE": {
        FileStyle.A: file_style_a_details_factory.create("ROE", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Return on Average Common Equity", "Financial Summary")
    },
    "ROIC": {
        FileStyle.A: file_style_a_details_factory.create("ROIC", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Return on Invested Capital", "Financial Summary")
    },
    "Current Ratio": {
        FileStyle.A: file_style_a_details_factory.create("Current Ratio", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Current Ratio", "Financial Summary")
    },
    "Quick Ratio": {
        FileStyle.A: file_style_a_details_factory.create("Quick Ratio", "Ratios - Key Metric"),
        FileStyle.B: file_style_b_details_factory.create("Quick Ratio", "Financial Summary")
    },
}
