from file_style import FileStyle, FileStyleDetails


file_style_configs_by_metric = {
    "ROA": {
        FileStyle.A: FileStyleDetails("A1", " | ", "%b-%Y", "Ratios - Key Metric", "Pretax ROA", "A1", "A3", "C6"),
        FileStyle.B: FileStyleDetails("B2", " (", "%d-%m-%Y", "Financial Summary", "Pretax ROA", "A1", "A14", "B15")
    },
    "Gross Margin": {
        FileStyle.A: FileStyleDetails("A1", " | ", "%b-%Y", "Ratios - Key Metric", "Pretax ROA", "A1", "A3", "C6"),
        FileStyle.B: FileStyleDetails("B2", " (", "%d-%m-%Y", "Financial Summary", "Pretax ROA", "A1", "A14", "B15")
    }
}
