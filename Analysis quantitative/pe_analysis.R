#!/usr/bin/env Rscript
# pe_analysis.R
# Comprehensive Private Equity Analysis Script
# Covers: Descriptive Statistics, Company Analysis, Trend Analysis,
#         Sectoral Analysis, Stage Analysis, Exit Analysis
# Usage: source("pe_analysis.R") in R or Rscript pe_analysis.R in terminal

## ============================================================================
## SECTION 1: PACKAGE SETUP
## ============================================================================

pkg_check <- function(pkgs) {
  to_install <- pkgs[!(pkgs %in% installed.packages()[, "Package"])]
  if (length(to_install)) {
    message("Installing packages: ", paste(to_install, collapse = ", "))
    install.packages(to_install, repos = "https://cran.rstudio.com/", quiet = TRUE)
  }
  invisible(lapply(pkgs, library, character.only = TRUE))
}

required_pkgs <- c(
  "readxl", "readr", "dplyr", "ggplot2", "tidyr", "lubridate",
  "corrplot", "reshape2", "scales", "RColorBrewer", "ggpubr",
  "plotly", "viridis", "gridExtra", "forcats"
)
pkg_check(required_pkgs)

## ============================================================================
## SECTION 2: HELPER FUNCTIONS
## ============================================================================

log_msg <- function(...) cat(format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "-", ..., "\n")

safe_read <- function(path) {
  if (!file.exists(path)) {
    return(NULL)
  }
  ext <- tolower(tools::file_ext(path))
  if (ext %in% c("xls", "xlsx")) {
    return(read_excel(path))
  }
  if (ext %in% c("csv", "txt")) {
    return(read_csv(path, show_col_types = FALSE))
  }
  return(NULL)
}

get_numeric_cols <- function(df) {
  names(df)[sapply(df, is.numeric)]
}

slugify <- function(x) {
  x <- tolower(gsub("[^a-zA-Z0-9]+", "_", x))
  gsub("_+", "_", gsub("^_|_$", "", x))
}

## ============================================================================
## SECTION 3: PLOT SAVING HELPERS
## ============================================================================

saved_images <- character(0)
saved_csvs <- character(0)

save_plot <- function(plot_obj, subdir, filename, width = 10, height = 6, dpi = 320) {
  dir_path <- file.path(output_dir, subdir)
  if (!dir.exists(dir_path)) dir.create(dir_path, recursive = TRUE)
  png_path <- file.path(dir_path, paste0(filename, ".png"))
  ggsave(png_path, plot_obj, width = width, height = height, dpi = dpi)
  saved_images <<- c(saved_images, png_path)
  log_msg("Saved chart:", subdir, "/", basename(png_path))
  invisible(png_path)
}

save_base_plot <- function(subdir, filename, plot_expr, width = 10, height = 6, res = 320) {
  dir_path <- file.path(output_dir, subdir)
  if (!dir.exists(dir_path)) dir.create(dir_path, recursive = TRUE)
  png_path <- file.path(dir_path, paste0(filename, ".png"))
  png(filename = png_path, width = width * res, height = height * res, res = res)
  on.exit(dev.off(), add = TRUE)
  plot_expr()
  saved_images <<- c(saved_images, png_path)
  log_msg("Saved chart:", subdir, "/", basename(png_path))
  invisible(png_path)
}

save_csv <- function(df, subdir, filename) {
  dir_path <- file.path(output_dir, subdir)
  if (!dir.exists(dir_path)) dir.create(dir_path, recursive = TRUE)
  csv_path <- file.path(dir_path, paste0(filename, ".csv"))
  write.csv(df, csv_path, row.names = FALSE)
  saved_csvs <<- c(saved_csvs, csv_path)
  log_msg("Saved CSV:", subdir, "/", basename(csv_path))
  invisible(csv_path)
}

## ============================================================================
## SECTION 4: CONFIGURATION
## ============================================================================

# Set base directory to current working directory where data files exist
base_dir <- "d:/office works/SST8654/Analysis quantitative"
investment_file <- file.path(base_dir, "Raw Dataset 2 ( PE Deal Value Investment).csv")
exit_file <- file.path(base_dir, "Raw Dataset 1 ( PE Deal value Exits.csv")
output_dir <- file.path(base_dir, "pe_analysis_outputs")

if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

log_msg("=== STARTING COMPREHENSIVE PE ANALYSIS ===")
log_msg("Output directory:", output_dir)

## ============================================================================
## SECTION 5: LOAD DATA
## ============================================================================

log_msg("Loading datasets...")
investment_data <- safe_read(investment_file)
exit_data <- safe_read(exit_file)

if (is.null(investment_data) && is.null(exit_data)) {
  stop("No datasets found. Check file paths:\n  - ", investment_file, "\n  - ", exit_file)
}

log_msg("Investment data:", ifelse(!is.null(investment_data), paste(nrow(investment_data), "rows"), "NOT FOUND"))
log_msg("Exit data:", ifelse(!is.null(exit_data), paste(nrow(exit_data), "rows"), "NOT FOUND"))

## ============================================================================
## SECTION 6: CLEAN COLUMN NAMES
## ============================================================================

clean_names <- function(df) {
  names(df) <- tolower(names(df))
  names(df) <- gsub("[^a-z0-9]+", "_", names(df))
  names(df) <- gsub("_+", "_", names(df))
  names(df) <- gsub("^_|_$", "", names(df))
  df
}

if (!is.null(investment_data)) investment_data <- clean_names(investment_data)
if (!is.null(exit_data)) exit_data <- clean_names(exit_data)

log_msg("Column names cleaned")

## ============================================================================
## SECTION 7: DATA CLEANING & PREPROCESSING
## ============================================================================

clean_numeric <- function(x) {
  as.numeric(gsub("[^0-9.-]", "", as.character(x)))
}

# Clean investment data
if (!is.null(investment_data)) {
  # Extract year from investment date
  date_cols <- grep("date|year", names(investment_data), ignore.case = TRUE, value = TRUE)
  if (length(date_cols) > 0) {
    date_col <- date_cols[1]
    investment_data$investment_year <- suppressWarnings(lubridate::year(lubridate::ymd(investment_data[[date_col]])))
  }

  # Identify key columns
  sector_cols <- grep("sector|industry", names(investment_data), ignore.case = TRUE, value = TRUE)
  stage_cols <- grep("stage|round", names(investment_data), ignore.case = TRUE, value = TRUE)
  company_cols <- grep("company|investee", names(investment_data), ignore.case = TRUE, value = TRUE)
  value_cols <- grep("value|equity|amount", names(investment_data), ignore.case = TRUE, value = TRUE)

  if (length(sector_cols) > 0) investment_data$sector <- investment_data[[sector_cols[1]]]
  if (length(stage_cols) > 0) investment_data$stage <- investment_data[[stage_cols[1]]]
  if (length(company_cols) > 0) investment_data$company <- investment_data[[company_cols[1]]]
  if (length(value_cols) > 0) {
    investment_data$deal_value <- clean_numeric(investment_data[[value_cols[1]]])
  }

  log_msg("Investment data preprocessed")
}

# Clean exit data
if (!is.null(exit_data)) {
  # Find date/year column
  date_cols <- grep("date|year|ann", names(exit_data), ignore.case = TRUE, value = TRUE)
  if (length(date_cols) > 0) {
    date_col <- date_cols[1]
    exit_data$exit_year <- suppressWarnings(lubridate::year(lubridate::ymd(exit_data[[date_col]])))
    if (all(is.na(exit_data$exit_year))) {
      # Try extracting year from numeric column
      exit_data$exit_year <- suppressWarnings(as.numeric(exit_data[[date_col]]))
    }
  }

  # Identify key columns
  company_cols <- grep("company|target|name", names(exit_data), ignore.case = TRUE, value = TRUE)
  value_cols <- grep("value|deal|amount", names(exit_data), ignore.case = TRUE, value = TRUE)
  sector_cols <- grep("sector|industry", names(exit_data), ignore.case = TRUE, value = TRUE)

  if (length(company_cols) > 0) exit_data$company <- exit_data[[company_cols[1]]]
  if (length(value_cols) > 0) {
    exit_data$exit_value <- clean_numeric(exit_data[[value_cols[1]]])
  }
  if (length(sector_cols) > 0) exit_data$sector <- exit_data[[sector_cols[1]]]

  log_msg("Exit data preprocessed")
}

## ============================================================================
## SECTION 4.1: DESCRIPTIVE STATISTICS
## ============================================================================

log_msg("=== SECTION 4.1: Descriptive Statistics ===")

# Investment Statistics
if (!is.null(investment_data)) {
  inv_stats <- data.frame(
    Metric = c(
      "Total Records", "Total Companies", "Total Sectors",
      "Year Range Start", "Year Range End",
      "Mean Deal Value (USD M)", "Median Deal Value (USD M)",
      "Max Deal Value (USD M)", "Min Deal Value (USD M)",
      "Total Deal Value (USD M)"
    ),
    Value = c(
      nrow(investment_data),
      length(unique(investment_data$company)),
      length(unique(investment_data$sector)),
      min(investment_data$investment_year, na.rm = TRUE),
      max(investment_data$investment_year, na.rm = TRUE),
      round(mean(investment_data$deal_value, na.rm = TRUE), 2),
      round(median(investment_data$deal_value, na.rm = TRUE), 2),
      round(max(investment_data$deal_value, na.rm = TRUE), 2),
      round(min(investment_data$deal_value, na.rm = TRUE), 2),
      round(sum(investment_data$deal_value, na.rm = TRUE), 2)
    )
  )
  save_csv(inv_stats, "statistics", "investment_descriptive_statistics")
}

# Exit Statistics
if (!is.null(exit_data)) {
  exit_stats <- data.frame(
    Metric = c(
      "Total Exit Records", "Total Companies",
      "Mean Exit Value (USD)", "Median Exit Value (USD)",
      "Max Exit Value (USD)", "Total Exit Value (USD)"
    ),
    Value = c(
      nrow(exit_data),
      length(unique(exit_data$company)),
      round(mean(exit_data$exit_value, na.rm = TRUE), 2),
      round(median(exit_data$exit_value, na.rm = TRUE), 2),
      round(max(exit_data$exit_value, na.rm = TRUE), 2),
      round(sum(exit_data$exit_value, na.rm = TRUE), 2)
    )
  )
  save_csv(exit_stats, "statistics", "exit_descriptive_statistics")
}

## ============================================================================
## SECTION 4.2: INVESTMENT SUMMARY - COMPANY WISE
## ============================================================================

log_msg("=== SECTION 4.2: Investment Summary (Company-wise) ===")

if (!is.null(investment_data) && "company" %in% names(investment_data)) {
  # Company-wise summary
  company_investment <- investment_data %>%
    filter(!is.na(company)) %>%
    group_by(company) %>%
    summarise(
      total_deals = n(),
      total_investment = sum(deal_value, na.rm = TRUE),
      avg_deal_size = mean(deal_value, na.rm = TRUE),
      first_investment = min(investment_year, na.rm = TRUE),
      last_investment = max(investment_year, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    arrange(desc(total_investment))

  save_csv(company_investment, "statistics", "company_investment_summary")

  # Top 20 Companies by Investment
  top_companies <- company_investment %>% slice_head(n = 20)

  p_company <- ggplot(top_companies, aes(x = reorder(company, total_investment), y = total_investment)) +
    geom_col(fill = "#2b8cbe", alpha = 0.85) +
    coord_flip() +
    scale_y_continuous(labels = scales::comma) +
    labs(
      title = "Top 20 Companies by Total Investment Received",
      subtitle = "PE Investment (USD Millions)",
      x = "Company",
      y = "Total Investment (USD Millions)"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.y = element_text(size = 8)
    )

  save_plot(p_company, "company_analysis", "chart_top20_companies_investment", width = 12, height = 10)

  # Deal Count by Company
  top_by_deals <- company_investment %>%
    arrange(desc(total_deals)) %>%
    slice_head(n = 20)

  p_deals <- ggplot(top_by_deals, aes(x = reorder(company, total_deals), y = total_deals)) +
    geom_col(fill = "#e41a1c", alpha = 0.85) +
    coord_flip() +
    labs(
      title = "Top 20 Companies by Number of Deals",
      subtitle = "Number of investment rounds received",
      x = "Company",
      y = "Number of Deals"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.y = element_text(size = 8)
    )

  save_plot(p_deals, "company_analysis", "chart_top20_companies_by_deal_count", width = 12, height = 10)
}

## ============================================================================
## SECTION 4.3: EXIT SUMMARY - COMPANY WISE
## ============================================================================

log_msg("=== SECTION 4.3: Exit Summary (Company-wise) ===")

if (!is.null(exit_data) && "company" %in% names(exit_data)) {
  # Company-wise exit summary
  company_exits <- exit_data %>%
    filter(!is.na(company)) %>%
    group_by(company) %>%
    summarise(
      total_exits = n(),
      total_exit_value = sum(exit_value, na.rm = TRUE),
      avg_exit_value = mean(exit_value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    arrange(desc(total_exit_value))

  save_csv(company_exits, "statistics", "company_exit_summary")

  # Top 20 Companies by Exit Value
  top_exits <- company_exits %>%
    filter(!is.na(total_exit_value) & total_exit_value > 0) %>%
    slice_head(n = 20)

  if (nrow(top_exits) > 0) {
    p_exit_company <- ggplot(top_exits, aes(x = reorder(company, total_exit_value), y = total_exit_value)) +
      geom_col(fill = "#1b9e77", alpha = 0.85) +
      coord_flip() +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Top 20 Companies by Exit Value",
        subtitle = "PE Exit Transactions",
        x = "Company",
        y = "Exit Value (USD)"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.y = element_text(size = 8)
      )

    save_plot(p_exit_company, "company_analysis", "chart_top20_companies_exit_value", width = 12, height = 10)
  }
}

## ============================================================================
## SECTION 4.4: TREND ANALYSIS
## ============================================================================

log_msg("=== SECTION 4.4: Trend Analysis ===")

if (!is.null(investment_data) && "investment_year" %in% names(investment_data)) {
  ## 4.4.1 Annual Deal Volume Trend
  annual_volume <- investment_data %>%
    filter(!is.na(investment_year)) %>%
    group_by(investment_year) %>%
    summarise(deal_count = n(), .groups = "drop") %>%
    arrange(investment_year)

  save_csv(annual_volume, "trend_analysis", "annual_deal_volume")

  p_volume <- ggplot(annual_volume, aes(x = investment_year, y = deal_count)) +
    geom_line(color = "#2b8cbe", size = 1.2) +
    geom_point(color = "#2b8cbe", size = 3) +
    geom_area(alpha = 0.2, fill = "#2b8cbe") +
    labs(
      title = "Annual Deal Volume Trend",
      subtitle = "Number of PE Investment Deals per Year",
      x = "Year",
      y = "Number of Deals"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )

  save_plot(p_volume, "trend_analysis", "chart_annual_deal_volume")

  ## 4.4.2 Annual Deal Value Trend
  annual_value <- investment_data %>%
    filter(!is.na(investment_year), !is.na(deal_value)) %>%
    group_by(investment_year) %>%
    summarise(
      total_value = sum(deal_value, na.rm = TRUE),
      avg_value = mean(deal_value, na.rm = TRUE),
      median_value = median(deal_value, na.rm = TRUE),
      deal_count = n(),
      .groups = "drop"
    ) %>%
    arrange(investment_year)

  save_csv(annual_value, "trend_analysis", "annual_deal_value")

  # Total Value Trend
  p_value <- ggplot(annual_value, aes(x = investment_year, y = total_value)) +
    geom_line(color = "#e41a1c", size = 1.2) +
    geom_point(color = "#e41a1c", size = 3) +
    geom_area(alpha = 0.2, fill = "#e41a1c") +
    scale_y_continuous(labels = scales::comma) +
    labs(
      title = "Annual Deal Value Trend",
      subtitle = "Total PE Investment Value per Year (USD Millions)",
      x = "Year",
      y = "Total Deal Value (USD Millions)"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )

  save_plot(p_value, "trend_analysis", "chart_annual_deal_value")

  # Average Deal Size Trend
  p_avg <- ggplot(annual_value, aes(x = investment_year, y = avg_value)) +
    geom_line(color = "#7570b3", size = 1.2) +
    geom_point(color = "#7570b3", size = 3) +
    scale_y_continuous(labels = scales::comma) +
    labs(
      title = "Average Deal Size Trend",
      subtitle = "Average PE Deal Value per Year (USD Millions)",
      x = "Year",
      y = "Average Deal Value (USD Millions)"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )

  save_plot(p_avg, "trend_analysis", "chart_average_deal_size")

  # Combined Volume and Value Chart
  annual_combined <- annual_value %>%
    mutate(value_scaled = total_value / max(total_value) * max(deal_count))

  p_combined <- ggplot(annual_combined, aes(x = investment_year)) +
    geom_bar(aes(y = deal_count), stat = "identity", fill = "#2b8cbe", alpha = 0.6) +
    geom_line(aes(y = value_scaled), color = "#e41a1c", size = 1.5) +
    geom_point(aes(y = value_scaled), color = "#e41a1c", size = 3) +
    scale_y_continuous(
      name = "Number of Deals",
      sec.axis = sec_axis(~ . / max(annual_combined$deal_count) * max(annual_value$total_value),
        name = "Total Value (USD M)", labels = scales::comma
      )
    ) +
    labs(
      title = "Deal Volume vs Deal Value Over Time",
      subtitle = "Bars = Deal Count, Line = Total Value",
      x = "Year"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1),
      axis.title.y.right = element_text(color = "#e41a1c"),
      axis.title.y.left = element_text(color = "#2b8cbe")
    )

  save_plot(p_combined, "trend_analysis", "chart_volume_vs_value_combined", width = 12)

  ## 4.4.3 Sectoral Allocations
  if ("sector" %in% names(investment_data)) {
    sector_summary <- investment_data %>%
      filter(!is.na(sector)) %>%
      group_by(sector) %>%
      summarise(
        deal_count = n(),
        total_value = sum(deal_value, na.rm = TRUE),
        avg_value = mean(deal_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(desc(total_value))

    save_csv(sector_summary, "sector_analysis", "sector_investment_summary")

    # Overall Sector Distribution Pie Chart
    top_sectors <- sector_summary %>% slice_head(n = 10)

    p_pie <- ggplot(top_sectors, aes(x = "", y = total_value, fill = sector)) +
      geom_bar(stat = "identity", width = 1) +
      coord_polar("y", start = 0) +
      scale_fill_brewer(palette = "Set3") +
      labs(
        title = "Sector Distribution by Investment Value",
        subtitle = "Top 10 Sectors",
        fill = "Sector"
      ) +
      theme_void() +
      theme(
        plot.title = element_text(face = "bold", size = 14, hjust = 0.5),
        legend.position = "right"
      )

    save_plot(p_pie, "sector_analysis", "chart_sector_distribution_pie", width = 12, height = 8)

    # Sector Bar Chart
    p_sector_bar <- ggplot(
      sector_summary %>% slice_head(n = 15),
      aes(x = reorder(sector, total_value), y = total_value)
    ) +
      geom_col(fill = "#66c2a5", alpha = 0.85) +
      coord_flip() +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Top 15 Sectors by Investment Value",
        x = "Sector",
        y = "Total Investment (USD Millions)"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_sector_bar, "sector_analysis", "chart_sector_investment_bar", width = 12, height = 8)

    ## 4.4.4 Sector Share Over Time
    sector_year <- investment_data %>%
      filter(!is.na(sector), !is.na(investment_year)) %>%
      group_by(investment_year, sector) %>%
      summarise(deal_count = n(), total_value = sum(deal_value, na.rm = TRUE), .groups = "drop")

    # Get top sectors for clarity
    top_sector_names <- sector_summary %>%
      slice_head(n = 8) %>%
      pull(sector)
    sector_year_top <- sector_year %>%
      mutate(sector = ifelse(sector %in% top_sector_names, sector, "Others")) %>%
      group_by(investment_year, sector) %>%
      summarise(deal_count = sum(deal_count), total_value = sum(total_value), .groups = "drop")

    save_csv(sector_year, "sector_analysis", "sector_year_breakdown")

    # Stacked Area Chart
    p_sector_stack <- ggplot(sector_year_top, aes(x = investment_year, y = total_value, fill = sector)) +
      geom_area(alpha = 0.8) +
      scale_fill_brewer(palette = "Set2") +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Sectoral Investment Allocation Over Time",
        subtitle = "Stacked area showing sector share by year",
        x = "Year",
        y = "Investment Value (USD Millions)",
        fill = "Sector"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_sector_stack, "sector_analysis", "chart_sector_allocation_stacked", width = 14, height = 8)

    # Sector Share Percentage Over Time
    sector_year_pct <- sector_year_top %>%
      group_by(investment_year) %>%
      mutate(pct = total_value / sum(total_value) * 100) %>%
      ungroup()

    p_sector_pct <- ggplot(sector_year_pct, aes(x = investment_year, y = pct, color = sector)) +
      geom_line(size = 1) +
      geom_point(size = 2) +
      scale_color_brewer(palette = "Set2") +
      labs(
        title = "Sector Share Percentage Over Time",
        x = "Year",
        y = "Share (%)",
        color = "Sector"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_sector_pct, "sector_analysis", "chart_sector_share_trend", width = 14, height = 8)

    ## 4.4.5 Sector Attractiveness Heatmap
    heatmap_data <- investment_data %>%
      filter(!is.na(investment_year), !is.na(sector)) %>%
      count(investment_year, sector, name = "count")

    heatmap_top <- heatmap_data %>%
      filter(sector %in% top_sector_names)

    p_heatmap <- ggplot(heatmap_top, aes(x = investment_year, y = sector, fill = count)) +
      geom_tile(color = "white") +
      scale_fill_gradientn(colors = c("#fff5f0", "#fee0d2", "#fcbba1", "#fc9272", "#fb6a4a", "#de2d26")) +
      labs(
        title = "Sector Attractiveness Heatmap",
        subtitle = "Deal count by Year and Sector",
        x = "Year",
        y = "Sector",
        fill = "Deal Count"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_heatmap, "sector_analysis", "chart_sector_heatmap", width = 14, height = 10)
  }
}

## ============================================================================
## SECTION 4.5: INVESTMENT STAGE ANALYSIS
## ============================================================================

log_msg("=== SECTION 4.5: Investment Stage Analysis ===")

if (!is.null(investment_data) && "stage" %in% names(investment_data)) {
  # Categorize stages
  investment_data <- investment_data %>%
    mutate(stage_category = case_when(
      grepl("seed|angel|early|pre-seed|startup", stage, ignore.case = TRUE) ~ "Early Stage",
      grepl("growth|expansion|series [c-z]|later", stage, ignore.case = TRUE) ~ "Growth Stage",
      grepl("buyout|late|acquisition|leveraged|lbo|mbo", stage, ignore.case = TRUE) ~ "Buyout/Late Stage",
      TRUE ~ "Other"
    ))

  ## 4.5.1 Distribution of Deal Stages
  stage_dist <- investment_data %>%
    filter(!is.na(stage_category)) %>%
    group_by(stage_category) %>%
    summarise(
      deal_count = n(),
      total_value = sum(deal_value, na.rm = TRUE),
      avg_value = mean(deal_value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(pct = deal_count / sum(deal_count) * 100)

  save_csv(stage_dist, "stage_analysis", "stage_distribution_summary")

  # Stage Pie Chart
  p_stage_pie <- ggplot(stage_dist, aes(x = "", y = deal_count, fill = stage_category)) +
    geom_bar(stat = "identity", width = 1) +
    coord_polar("y", start = 0) +
    scale_fill_manual(values = c(
      "Early Stage" = "#4daf4a", "Growth Stage" = "#377eb8",
      "Buyout/Late Stage" = "#e41a1c", "Other" = "#999999"
    )) +
    labs(
      title = "Distribution of Deal Stages",
      subtitle = "Early Stage vs Growth vs Buyout",
      fill = "Stage"
    ) +
    theme_void() +
    theme(
      plot.title = element_text(face = "bold", size = 14, hjust = 0.5)
    )

  save_plot(p_stage_pie, "stage_analysis", "chart_stage_distribution_pie", width = 10, height = 8)

  # Stage Bar Chart
  p_stage_bar <- ggplot(stage_dist, aes(x = reorder(stage_category, deal_count), y = deal_count, fill = stage_category)) +
    geom_col(alpha = 0.85) +
    geom_text(aes(label = paste0(round(pct, 1), "%")), hjust = -0.1, size = 4) +
    coord_flip() +
    scale_fill_manual(values = c(
      "Early Stage" = "#4daf4a", "Growth Stage" = "#377eb8",
      "Buyout/Late Stage" = "#e41a1c", "Other" = "#999999"
    )) +
    labs(
      title = "Deal Count by Investment Stage",
      x = "Stage",
      y = "Number of Deals"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      legend.position = "none"
    ) +
    expand_limits(y = max(stage_dist$deal_count) * 1.15)

  save_plot(p_stage_bar, "stage_analysis", "chart_stage_distribution_bar")

  ## 4.5.2 Trend in Control vs Minority (Buyout trend after 2016 IBC)
  stage_year <- investment_data %>%
    filter(!is.na(investment_year), !is.na(stage_category)) %>%
    group_by(investment_year, stage_category) %>%
    summarise(deal_count = n(), .groups = "drop") %>%
    group_by(investment_year) %>%
    mutate(pct = deal_count / sum(deal_count) * 100) %>%
    ungroup()

  save_csv(stage_year, "stage_analysis", "stage_trend_by_year")

  # Focus on Buyout trend
  buyout_trend <- stage_year %>%
    filter(stage_category == "Buyout/Late Stage")

  p_buyout <- ggplot(buyout_trend, aes(x = investment_year, y = pct)) +
    geom_line(color = "#e41a1c", size = 1.2) +
    geom_point(color = "#e41a1c", size = 3) +
    geom_vline(xintercept = 2016, linetype = "dashed", color = "gray40", size = 1) +
    annotate("text",
      x = 2016.5, y = max(buyout_trend$pct, na.rm = TRUE) * 0.9,
      label = "IBC 2016", hjust = 0, fontface = "italic"
    ) +
    labs(
      title = "Buyout/Late Stage Investment Trend",
      subtitle = "Percentage of deals that are buyouts (Note: IBC 2016 marked)",
      x = "Year",
      y = "Buyout Share (%)"
    ) +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14))

  save_plot(p_buyout, "stage_analysis", "chart_buyout_trend_ibc", width = 12)

  ## 4.5.3 Stage Stacked Area Chart
  p_stage_stack <- ggplot(stage_year, aes(x = investment_year, y = deal_count, fill = stage_category)) +
    geom_area(alpha = 0.8, position = "stack") +
    scale_fill_manual(values = c(
      "Early Stage" = "#4daf4a", "Growth Stage" = "#377eb8",
      "Buyout/Late Stage" = "#e41a1c", "Other" = "#999999"
    )) +
    geom_vline(xintercept = 2016, linetype = "dashed", color = "gray40") +
    labs(
      title = "Investment Stage Composition Over Time",
      subtitle = "Stacked area chart showing stage evolution",
      x = "Year",
      y = "Number of Deals",
      fill = "Stage"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )

  save_plot(p_stage_stack, "stage_analysis", "chart_stage_trend_stacked", width = 14, height = 8)
}

## ============================================================================
## SECTION 4.6: ANALYSIS OF PRIVATE EQUITY EXITS
## ============================================================================

log_msg("=== SECTION 4.6: PE Exit Analysis ===")

if (!is.null(exit_data)) {
  ## 4.6.1 Annual Exit Activity
  if ("exit_year" %in% names(exit_data)) {
    annual_exits <- exit_data %>%
      filter(!is.na(exit_year)) %>%
      group_by(exit_year) %>%
      summarise(
        exit_count = n(),
        total_exit_value = sum(exit_value, na.rm = TRUE),
        avg_exit_value = mean(exit_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      filter(exit_year >= 2000 & exit_year <= 2025) %>%
      arrange(exit_year)

    save_csv(annual_exits, "exit_analysis", "annual_exit_activity")

    # Exit Count Trend
    p_exit_count <- ggplot(annual_exits, aes(x = exit_year, y = exit_count)) +
      geom_line(color = "#1b9e77", size = 1.2) +
      geom_point(color = "#1b9e77", size = 3) +
      geom_area(alpha = 0.2, fill = "#1b9e77") +
      labs(
        title = "Annual PE Exit Activity (2005-2024)",
        subtitle = "Number of exits per year",
        x = "Year",
        y = "Number of Exits"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_exit_count, "exit_analysis", "chart_annual_exit_count")

    # Exit Value Trend
    p_exit_value <- ggplot(annual_exits, aes(x = exit_year, y = total_exit_value)) +
      geom_line(color = "#d95f02", size = 1.2) +
      geom_point(color = "#d95f02", size = 3) +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Annual PE Exit Value Trend",
        subtitle = "Total exit value per year (USD)",
        x = "Year",
        y = "Total Exit Value (USD)"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_exit_value, "exit_analysis", "chart_annual_exit_value")

    # Average Exit Value
    p_exit_avg <- ggplot(annual_exits, aes(x = exit_year, y = avg_exit_value)) +
      geom_line(color = "#7570b3", size = 1.2) +
      geom_point(color = "#7570b3", size = 3) +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Average Exit Deal Size Over Time",
        x = "Year",
        y = "Average Exit Value (USD)"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_exit_avg, "exit_analysis", "chart_average_exit_value")
  }

  ## 4.6.2 Exit Types Distribution (if exit_type column exists)
  exit_type_cols <- grep("type|mode|exit|route", names(exit_data), ignore.case = TRUE, value = TRUE)
  if (length(exit_type_cols) > 0) {
    exit_data$exit_type <- exit_data[[exit_type_cols[1]]]

    exit_type_dist <- exit_data %>%
      filter(!is.na(exit_type)) %>%
      group_by(exit_type) %>%
      summarise(
        count = n(),
        total_value = sum(exit_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(pct = count / sum(count) * 100) %>%
      arrange(desc(count))

    save_csv(exit_type_dist, "exit_analysis", "exit_type_distribution")

    # Exit Type Bar Chart
    p_exit_type <- ggplot(
      exit_type_dist %>% slice_head(n = 10),
      aes(x = reorder(exit_type, count), y = count, fill = exit_type)
    ) +
      geom_col(alpha = 0.85) +
      coord_flip() +
      scale_fill_brewer(palette = "Set2") +
      labs(
        title = "Exit Types Distribution",
        subtitle = "IPO, M&A, Secondary Sale, Buyback, etc.",
        x = "Exit Type",
        y = "Number of Exits"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        legend.position = "none"
      )

    save_plot(p_exit_type, "exit_analysis", "chart_exit_types_bar", width = 12)

    # Exit Type Pie Chart
    p_exit_type_pie <- ggplot(
      exit_type_dist %>% slice_head(n = 6),
      aes(x = "", y = count, fill = exit_type)
    ) +
      geom_bar(stat = "identity", width = 1) +
      coord_polar("y", start = 0) +
      scale_fill_brewer(palette = "Set2") +
      labs(
        title = "Exit Type Proportions",
        fill = "Exit Type"
      ) +
      theme_void() +
      theme(
        plot.title = element_text(face = "bold", size = 14, hjust = 0.5)
      )

    save_plot(p_exit_type_pie, "exit_analysis", "chart_exit_types_pie")
  }

  ## 4.6.3 Sector-Wise Exit Patterns
  if ("sector" %in% names(exit_data)) {
    sector_exits <- exit_data %>%
      filter(!is.na(sector)) %>%
      group_by(sector) %>%
      summarise(
        exit_count = n(),
        total_exit_value = sum(exit_value, na.rm = TRUE),
        avg_exit_value = mean(exit_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(desc(exit_count))

    save_csv(sector_exits, "exit_analysis", "sector_exit_summary")

    # Exits by Sector
    p_sector_exit <- ggplot(
      sector_exits %>% slice_head(n = 15),
      aes(x = reorder(sector, exit_count), y = exit_count)
    ) +
      geom_col(fill = "#b2df8a", alpha = 0.85) +
      coord_flip() +
      labs(
        title = "Exit Count by Sector",
        subtitle = "Which sectors produce the most exits?",
        x = "Sector",
        y = "Number of Exits"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_sector_exit, "exit_analysis", "chart_exit_by_sector", width = 12, height = 8)

    # Exit Value by Sector (Liquidity)
    p_sector_liquidity <- ggplot(
      sector_exits %>%
        filter(!is.na(total_exit_value) & total_exit_value > 0) %>%
        slice_head(n = 15),
      aes(x = reorder(sector, total_exit_value), y = total_exit_value)
    ) +
      geom_col(fill = "#33a02c", alpha = 0.85) +
      coord_flip() +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Exit Value by Sector",
        subtitle = "Which sectors offer highest liquidity?",
        x = "Sector",
        y = "Total Exit Value (USD)"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_sector_liquidity, "exit_analysis", "chart_exit_value_by_sector", width = 12, height = 8)
  }
}

## ============================================================================
## SECTION 4.7: EXIT PERFORMANCE ANALYSIS
## ============================================================================

log_msg("=== SECTION 4.7: Exit Performance Analysis ===")

if (!is.null(exit_data)) {
  ## 4.7.1 Exit Value Distribution
  if ("exit_value" %in% names(exit_data)) {
    exit_values_clean <- exit_data %>%
      filter(!is.na(exit_value), exit_value > 0)

    # Histogram of Exit Values
    p_exit_dist <- ggplot(exit_values_clean, aes(x = exit_value)) +
      geom_histogram(bins = 50, fill = "#7570b3", color = "white", alpha = 0.8) +
      scale_x_continuous(labels = scales::comma) +
      labs(
        title = "Exit Value Distribution",
        subtitle = "Distribution of PE exit deal values",
        x = "Exit Value (USD)",
        y = "Frequency"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_exit_dist, "exit_analysis", "chart_exit_value_distribution", width = 12)

    # Log-scale histogram for better visualization
    p_exit_dist_log <- ggplot(exit_values_clean, aes(x = exit_value)) +
      geom_histogram(bins = 50, fill = "#e7298a", color = "white", alpha = 0.8) +
      scale_x_log10(labels = scales::comma) +
      labs(
        title = "Exit Value Distribution (Log Scale)",
        subtitle = "Logarithmic scale for better visibility of distribution",
        x = "Exit Value (USD, log scale)",
        y = "Frequency"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_exit_dist_log, "exit_analysis", "chart_exit_value_distribution_log", width = 12)
  }

  ## 4.7.2 Exit Value by Sector (if sector exists)
  if (all(c("sector", "exit_value") %in% names(exit_data))) {
    sector_exit_values <- exit_data %>%
      filter(!is.na(sector), !is.na(exit_value), exit_value > 0) %>%
      group_by(sector) %>%
      filter(n() >= 3) %>%
      ungroup()

    if (nrow(sector_exit_values) > 0) {
      # Get top 10 sectors by count
      top_sectors <- sector_exit_values %>%
        count(sector, sort = TRUE) %>%
        slice_head(n = 10) %>%
        pull(sector)

      sector_exit_filtered <- sector_exit_values %>%
        filter(sector %in% top_sectors)

      # Boxplot by Sector
      p_exit_sector_box <- ggplot(sector_exit_filtered, aes(
        x = reorder(sector, exit_value, FUN = median),
        y = exit_value, fill = sector
      )) +
        geom_boxplot(alpha = 0.7, outlier.shape = 21) +
        coord_flip() +
        scale_y_log10(labels = scales::comma) +
        scale_fill_brewer(palette = "Set3") +
        labs(
          title = "Exit Value Distribution by Sector",
          subtitle = "Boxplot showing median and spread (log scale)",
          x = "Sector",
          y = "Exit Value (USD, log scale)"
        ) +
        theme_minimal() +
        theme(
          plot.title = element_text(face = "bold", size = 14),
          legend.position = "none"
        )

      save_plot(p_exit_sector_box, "exit_analysis", "chart_exit_value_boxplot_sector", width = 12, height = 10)
    }
  }
}

## ============================================================================
## SECTION 8: SAVE CLEANED DATA
## ============================================================================

log_msg("=== Saving Cleaned Datasets ===")

if (!is.null(investment_data)) {
  save_csv(investment_data, "cleaned_data", "investment_data_cleaned")
}

if (!is.null(exit_data)) {
  save_csv(exit_data, "cleaned_data", "exit_data_cleaned")
}

## ============================================================================
## SECTION 9: CORRELATION ANALYSIS
## ============================================================================

log_msg("=== Additional Analysis: Correlations ===")

if (!is.null(investment_data)) {
  numeric_cols <- get_numeric_cols(investment_data)
  numeric_cols <- numeric_cols[!numeric_cols %in% c("investment_year", "round_number")]

  if (length(numeric_cols) >= 2) {
    corr_mat <- tryCatch(cor(investment_data[numeric_cols], use = "pairwise.complete.obs"), error = function(e) NULL)
    if (!is.null(corr_mat) && !any(is.na(corr_mat))) {
      save_base_plot(
        "statistics", "chart_correlation_matrix",
        plot_expr = function() {
          corrplot::corrplot(corr_mat,
            method = "color", type = "upper",
            tl.cex = 0.8, mar = c(0, 0, 2, 0),
            title = "Correlation Matrix - Investment Variables"
          )
        },
        width = 12,
        height = 10
      )
    }
  }
}

## ============================================================================
## SECTION 4.8: DEAL-LEVEL ANALYSIS
## ============================================================================

log_msg("=== SECTION 4.8: Deal-Level Analysis ===")

if (!is.null(investment_data)) {
  # Ensure we have date column for deal-level analysis
  date_cols <- grep("date|investment_date", names(investment_data), ignore.case = TRUE, value = TRUE)
  if (length(date_cols) > 0) {
    date_col <- date_cols[1]
    investment_data$deal_date <- suppressWarnings(lubridate::ymd(investment_data[[date_col]]))
    investment_data$deal_month <- lubridate::month(investment_data$deal_date)
    investment_data$deal_quarter <- lubridate::quarter(investment_data$deal_date)
  }

  ## 4.8.1 Individual Deal Summary
  deal_summary <- investment_data %>%
    filter(!is.na(company), !is.na(deal_value)) %>%
    mutate(
      deal_id = row_number(),
      deal_size_category = case_when(
        deal_value < 10 ~ "Micro (<$10M)",
        deal_value < 50 ~ "Small ($10-50M)",
        deal_value < 100 ~ "Medium ($50-100M)",
        deal_value < 500 ~ "Large ($100-500M)",
        TRUE ~ "Mega (>$500M)"
      )
    ) %>%
    select(
      deal_id, company, sector, stage, investment_year, deal_date,
      deal_value, deal_size_category, everything()
    )

  save_csv(
    deal_summary %>% select(
      deal_id, company, sector, stage, investment_year,
      deal_value, deal_size_category
    ),
    "deal_analysis", "deal_level_summary"
  )

  ## 4.8.2 Deal Size Distribution by Category
  deal_size_dist <- deal_summary %>%
    group_by(deal_size_category) %>%
    summarise(
      deal_count = n(),
      total_value = sum(deal_value, na.rm = TRUE),
      avg_value = mean(deal_value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      deal_size_category = factor(deal_size_category, levels = c(
        "Micro (<$10M)", "Small ($10-50M)", "Medium ($50-100M)",
        "Large ($100-500M)", "Mega (>$500M)"
      )),
      pct_count = deal_count / sum(deal_count) * 100,
      pct_value = total_value / sum(total_value) * 100
    ) %>%
    arrange(deal_size_category)

  save_csv(deal_size_dist, "deal_analysis", "deal_size_distribution")

  # Deal Size Bar Chart (by count)
  p_deal_size_count <- ggplot(deal_size_dist, aes(x = deal_size_category, y = deal_count, fill = deal_size_category)) +
    geom_col(alpha = 0.85) +
    geom_text(aes(label = paste0(round(pct_count, 1), "%")), vjust = -0.3, size = 4) +
    scale_fill_brewer(palette = "RdYlBu", direction = -1) +
    labs(
      title = "Deal Size Distribution (by Count)",
      subtitle = "Number of deals in each size category",
      x = "Deal Size Category",
      y = "Number of Deals"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 30, hjust = 1),
      legend.position = "none"
    ) +
    expand_limits(y = max(deal_size_dist$deal_count) * 1.1)

  save_plot(p_deal_size_count, "deal_analysis", "chart_deal_size_by_count", width = 12)

  # Deal Size Bar Chart (by value)
  p_deal_size_value <- ggplot(deal_size_dist, aes(x = deal_size_category, y = total_value, fill = deal_size_category)) +
    geom_col(alpha = 0.85) +
    geom_text(aes(label = paste0(round(pct_value, 1), "%")), vjust = -0.3, size = 4) +
    scale_fill_brewer(palette = "RdYlBu", direction = -1) +
    scale_y_continuous(labels = scales::comma) +
    labs(
      title = "Deal Size Distribution (by Value)",
      subtitle = "Total investment value in each size category",
      x = "Deal Size Category",
      y = "Total Deal Value (USD Millions)"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 30, hjust = 1),
      legend.position = "none"
    ) +
    expand_limits(y = max(deal_size_dist$total_value) * 1.1)

  save_plot(p_deal_size_value, "deal_analysis", "chart_deal_size_by_value", width = 12)

  # Deal Size Pie Chart
  p_deal_size_pie <- ggplot(deal_size_dist, aes(x = "", y = deal_count, fill = deal_size_category)) +
    geom_bar(stat = "identity", width = 1) +
    coord_polar("y", start = 0) +
    scale_fill_brewer(palette = "RdYlBu", direction = -1) +
    labs(
      title = "Deal Size Distribution",
      fill = "Deal Size"
    ) +
    theme_void() +
    theme(plot.title = element_text(face = "bold", size = 14, hjust = 0.5))

  save_plot(p_deal_size_pie, "deal_analysis", "chart_deal_size_pie")

  ## 4.8.3 Deal Size Trend Over Time
  deal_size_trend <- deal_summary %>%
    filter(!is.na(investment_year), !is.na(deal_size_category)) %>%
    group_by(investment_year, deal_size_category) %>%
    summarise(deal_count = n(), .groups = "drop") %>%
    mutate(deal_size_category = factor(deal_size_category, levels = c(
      "Micro (<$10M)", "Small ($10-50M)", "Medium ($50-100M)",
      "Large ($100-500M)", "Mega (>$500M)"
    )))

  save_csv(deal_size_trend, "deal_analysis", "deal_size_trend_by_year")

  # Stacked Area: Deal Size Over Time
  p_deal_size_stack <- ggplot(deal_size_trend, aes(x = investment_year, y = deal_count, fill = deal_size_category)) +
    geom_area(alpha = 0.8, position = "stack") +
    scale_fill_brewer(palette = "RdYlBu", direction = -1) +
    labs(
      title = "Deal Size Composition Over Time",
      subtitle = "How deal sizes have evolved across years",
      x = "Year",
      y = "Number of Deals",
      fill = "Deal Size"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(face = "bold", size = 14),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )

  save_plot(p_deal_size_stack, "deal_analysis", "chart_deal_size_trend_stacked", width = 14, height = 8)

  ## 4.8.4 Quarterly Deal Activity
  if ("deal_quarter" %in% names(deal_summary)) {
    quarterly_deals <- deal_summary %>%
      filter(!is.na(investment_year), !is.na(deal_quarter)) %>%
      mutate(year_quarter = paste0(investment_year, "-Q", deal_quarter)) %>%
      group_by(year_quarter, investment_year, deal_quarter) %>%
      summarise(
        deal_count = n(),
        total_value = sum(deal_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(investment_year, deal_quarter)

    save_csv(quarterly_deals, "deal_analysis", "quarterly_deal_activity")

    # Filter to last 5 years for clarity
    max_year <- max(quarterly_deals$investment_year, na.rm = TRUE)
    quarterly_recent <- quarterly_deals %>% filter(investment_year >= max_year - 5)

    if (nrow(quarterly_recent) > 4) {
      p_quarterly <- ggplot(quarterly_recent, aes(x = year_quarter, y = deal_count)) +
        geom_col(fill = "#2b8cbe", alpha = 0.8) +
        geom_line(aes(group = 1), color = "#e41a1c", size = 1) +
        labs(
          title = "Quarterly Deal Activity (Last 5 Years)",
          subtitle = "Number of deals per quarter",
          x = "Quarter",
          y = "Number of Deals"
        ) +
        theme_minimal() +
        theme(
          plot.title = element_text(face = "bold", size = 14),
          axis.text.x = element_text(angle = 90, hjust = 1, size = 7)
        )

      save_plot(p_quarterly, "deal_analysis", "chart_quarterly_deal_activity", width = 16, height = 8)
    }
  }

  ## 4.8.5 Monthly Seasonality Analysis
  if ("deal_month" %in% names(deal_summary)) {
    monthly_pattern <- deal_summary %>%
      filter(!is.na(deal_month)) %>%
      group_by(deal_month) %>%
      summarise(
        deal_count = n(),
        total_value = sum(deal_value, na.rm = TRUE),
        avg_value = mean(deal_value, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      mutate(month_name = month.abb[deal_month])

    save_csv(monthly_pattern, "deal_analysis", "monthly_seasonality")

    p_monthly <- ggplot(monthly_pattern, aes(x = factor(deal_month), y = deal_count)) +
      geom_col(fill = "#66c2a5", alpha = 0.85) +
      geom_line(aes(group = 1), color = "#e41a1c", size = 1) +
      geom_point(color = "#e41a1c", size = 3) +
      scale_x_discrete(labels = month.abb) +
      labs(
        title = "Monthly Seasonality Pattern",
        subtitle = "Which months see most PE activity?",
        x = "Month",
        y = "Number of Deals"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_monthly, "deal_analysis", "chart_monthly_seasonality")
  }

  ## 4.8.6 Investor-Level Analysis
  investor_cols <- grep("investor|firm", names(investment_data), ignore.case = TRUE, value = TRUE)
  if (length(investor_cols) > 0) {
    investor_col <- investor_cols[1]

    investor_summary <- investment_data %>%
      rename(investor = !!investor_col) %>%
      filter(!is.na(investor)) %>%
      group_by(investor) %>%
      summarise(
        deal_count = n(),
        total_invested = sum(deal_value, na.rm = TRUE),
        avg_deal_size = mean(deal_value, na.rm = TRUE),
        companies_invested = n_distinct(company),
        sectors_covered = n_distinct(sector),
        first_deal_year = min(investment_year, na.rm = TRUE),
        last_deal_year = max(investment_year, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      arrange(desc(total_invested))

    save_csv(investor_summary, "deal_analysis", "investor_level_summary")

    # Top 20 Investors by Investment Value
    top_investors <- investor_summary %>% slice_head(n = 20)

    p_investors <- ggplot(top_investors, aes(x = reorder(investor, total_invested), y = total_invested)) +
      geom_col(fill = "#7570b3", alpha = 0.85) +
      coord_flip() +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Top 20 PE Investors by Total Investment",
        subtitle = "Firms with highest cumulative investment value",
        x = "Investor",
        y = "Total Investment (USD Millions)"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.y = element_text(size = 8)
      )

    save_plot(p_investors, "deal_analysis", "chart_top20_investors", width = 14, height = 10)

    # Investors by Deal Count
    top_investors_count <- investor_summary %>%
      arrange(desc(deal_count)) %>%
      slice_head(n = 20)

    p_investors_count <- ggplot(top_investors_count, aes(x = reorder(investor, deal_count), y = deal_count)) +
      geom_col(fill = "#e7298a", alpha = 0.85) +
      coord_flip() +
      labs(
        title = "Top 20 Most Active PE Investors",
        subtitle = "Firms with highest number of deals",
        x = "Investor",
        y = "Number of Deals"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.y = element_text(size = 8)
      )

    save_plot(p_investors_count, "deal_analysis", "chart_top20_investors_by_deals", width = 14, height = 10)
  }

  ## 4.8.7 Deal Value Distribution Histogram
  p_deal_hist <- ggplot(
    deal_summary %>% filter(!is.na(deal_value), deal_value > 0),
    aes(x = deal_value)
  ) +
    geom_histogram(bins = 50, fill = "#2b8cbe", color = "white", alpha = 0.8) +
    scale_x_continuous(labels = scales::comma) +
    labs(
      title = "Deal Value Distribution",
      subtitle = "Histogram of individual deal values",
      x = "Deal Value (USD Millions)",
      y = "Frequency"
    ) +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14))

  save_plot(p_deal_hist, "deal_analysis", "chart_deal_value_histogram")

  # Log-scale histogram
  p_deal_hist_log <- ggplot(
    deal_summary %>% filter(!is.na(deal_value), deal_value > 0),
    aes(x = deal_value)
  ) +
    geom_histogram(bins = 50, fill = "#e41a1c", color = "white", alpha = 0.8) +
    scale_x_log10(labels = scales::comma) +
    labs(
      title = "Deal Value Distribution (Log Scale)",
      subtitle = "Logarithmic scale for better visibility",
      x = "Deal Value (USD Millions, log scale)",
      y = "Frequency"
    ) +
    theme_minimal() +
    theme(plot.title = element_text(face = "bold", size = 14))

  save_plot(p_deal_hist_log, "deal_analysis", "chart_deal_value_histogram_log")

  ## 4.8.8 Deal Value by Sector Boxplot
  if ("sector" %in% names(deal_summary)) {
    top_sectors <- deal_summary %>%
      filter(!is.na(sector)) %>%
      count(sector, sort = TRUE) %>%
      slice_head(n = 12) %>%
      pull(sector)

    deal_sector_box <- deal_summary %>%
      filter(sector %in% top_sectors, !is.na(deal_value), deal_value > 0)

    if (nrow(deal_sector_box) > 10) {
      p_sector_box <- ggplot(deal_sector_box, aes(
        x = reorder(sector, deal_value, FUN = median),
        y = deal_value, fill = sector
      )) +
        geom_boxplot(alpha = 0.7, outlier.shape = 21) +
        coord_flip() +
        scale_y_log10(labels = scales::comma) +
        scale_fill_brewer(palette = "Set3") +
        labs(
          title = "Deal Value Distribution by Sector",
          subtitle = "Boxplot showing spread and median (log scale)",
          x = "Sector",
          y = "Deal Value (USD Millions, log scale)"
        ) +
        theme_minimal() +
        theme(
          plot.title = element_text(face = "bold", size = 14),
          legend.position = "none"
        )

      save_plot(p_sector_box, "deal_analysis", "chart_deal_value_by_sector_boxplot", width = 12, height = 10)
    }
  }

  ## 4.8.9 Deal Value by Stage Boxplot
  if ("stage_category" %in% names(deal_summary)) {
    stage_box_data <- deal_summary %>%
      filter(!is.na(stage_category), !is.na(deal_value), deal_value > 0)

    if (nrow(stage_box_data) > 10) {
      p_stage_box <- ggplot(stage_box_data, aes(x = stage_category, y = deal_value, fill = stage_category)) +
        geom_boxplot(alpha = 0.7, outlier.shape = 21) +
        scale_y_log10(labels = scales::comma) +
        scale_fill_manual(values = c(
          "Early Stage" = "#4daf4a", "Growth Stage" = "#377eb8",
          "Buyout/Late Stage" = "#e41a1c", "Other" = "#999999"
        )) +
        labs(
          title = "Deal Value Distribution by Investment Stage",
          subtitle = "Buyouts typically have higher deal values",
          x = "Investment Stage",
          y = "Deal Value (USD Millions, log scale)"
        ) +
        theme_minimal() +
        theme(
          plot.title = element_text(face = "bold", size = 14),
          legend.position = "none"
        )

      save_plot(p_stage_box, "deal_analysis", "chart_deal_value_by_stage_boxplot")
    }
  }

  ## 4.8.10 Average Deal Size Trend by Sector
  if ("sector" %in% names(deal_summary)) {
    top_sectors <- deal_summary %>%
      count(sector, sort = TRUE) %>%
      slice_head(n = 6) %>%
      pull(sector)

    sector_deal_trend <- deal_summary %>%
      filter(sector %in% top_sectors, !is.na(investment_year), !is.na(deal_value)) %>%
      group_by(investment_year, sector) %>%
      summarise(
        avg_deal_value = mean(deal_value, na.rm = TRUE),
        deal_count = n(),
        .groups = "drop"
      )

    save_csv(sector_deal_trend, "deal_analysis", "sector_avg_deal_trend")

    p_sector_trend <- ggplot(sector_deal_trend, aes(
      x = investment_year, y = avg_deal_value,
      color = sector, group = sector
    )) +
      geom_line(size = 1.2) +
      geom_point(size = 2) +
      scale_y_continuous(labels = scales::comma) +
      scale_color_brewer(palette = "Set2") +
      labs(
        title = "Average Deal Size Trend by Sector",
        subtitle = "Top 6 sectors - how average deal sizes evolved",
        x = "Year",
        y = "Average Deal Value (USD Millions)",
        color = "Sector"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )

    save_plot(p_sector_trend, "deal_analysis", "chart_avg_deal_size_by_sector", width = 14, height = 8)
  }

  ## 4.8.11 Round Number Analysis
  round_cols <- grep("round|number", names(investment_data), ignore.case = TRUE, value = TRUE)
  if (length(round_cols) > 0) {
    round_col <- round_cols[1]

    round_analysis <- investment_data %>%
      rename(round_num = !!round_col) %>%
      filter(!is.na(round_num)) %>%
      mutate(round_num = as.numeric(round_num)) %>%
      filter(!is.na(round_num), round_num > 0, round_num <= 20) %>%
      group_by(round_num) %>%
      summarise(
        deal_count = n(),
        total_value = sum(deal_value, na.rm = TRUE),
        avg_value = mean(deal_value, na.rm = TRUE),
        .groups = "drop"
      )

    save_csv(round_analysis, "deal_analysis", "round_number_analysis")

    # Round Number Bar Chart
    p_round <- ggplot(round_analysis, aes(x = factor(round_num), y = deal_count)) +
      geom_col(fill = "#66c2a5", alpha = 0.85) +
      labs(
        title = "Distribution of Investment Rounds",
        subtitle = "Number of deals by round number (Series A=1, B=2, etc.)",
        x = "Round Number",
        y = "Number of Deals"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_round, "deal_analysis", "chart_round_distribution")

    # Average Deal Value by Round
    p_round_value <- ggplot(round_analysis, aes(x = factor(round_num), y = avg_value)) +
      geom_col(fill = "#7570b3", alpha = 0.85) +
      geom_line(aes(group = 1), color = "#e41a1c", size = 1) +
      geom_point(color = "#e41a1c", size = 3) +
      scale_y_continuous(labels = scales::comma) +
      labs(
        title = "Average Deal Value by Round Number",
        subtitle = "Later rounds typically have larger deal sizes",
        x = "Round Number",
        y = "Average Deal Value (USD Millions)"
      ) +
      theme_minimal() +
      theme(plot.title = element_text(face = "bold", size = 14))

    save_plot(p_round_value, "deal_analysis", "chart_avg_value_by_round")
  }

  log_msg("Deal-level analysis complete - multiple outputs saved to deal_analysis/")
}

## ============================================================================
## SECTION 10: FINAL SUMMARY REPORT
## ============================================================================

log_msg("")
log_msg("============================================")
log_msg("=== COMPREHENSIVE PE ANALYSIS COMPLETE ===")
log_msg("============================================")
log_msg("")
log_msg("Output Directory:", output_dir)
log_msg("")
log_msg("CSV FILES SAVED:", length(saved_csvs))
for (f in saved_csvs) log_msg("  -", basename(f))
log_msg("")
log_msg("CHARTS SAVED:", length(saved_images))
for (f in saved_images) log_msg("  -", basename(f))
log_msg("")
log_msg("Analysis covers:")
log_msg("  - Section 4.1: Descriptive Statistics")
log_msg("  - Section 4.2: Company-wise Investment Summary")
log_msg("  - Section 4.3: Company-wise Exit Summary")
log_msg("  - Section 4.4: Trend Analysis (Volume, Value, Sector)")
log_msg("  - Section 4.5: Investment Stage Analysis (Early/Growth/Buyout)")
log_msg("  - Section 4.6: PE Exit Analysis")
log_msg("  - Section 4.7: Exit Performance Analysis")
log_msg("  - Section 4.8: Deal-Level Analysis (Size, Quarterly, Monthly, Investor, Round)")
log_msg("")
log_msg("Open the output folder to view all results:")
log_msg("  ", output_dir)
log_msg("")
