# =============================================================================
# Lab Tracker — read-only Shiny front end for Lab_Tracker_2026.xlsx
#
# Key: 2D BARCODE. Every view is built from one table with one row per barcode,
# assembled from ALL_FinalReports (barcode_2d), GSA, WGS and Awaiting_GSA, plus
# two optional workbooks: BackupInventoryDatabase (plate/position of each backup
# tube) and OLD_FinalReports (older biorepository final reports).
# Original List (SampleID-keyed) is bridged in via the barcode's SampleID,
# preferring the Original List row whose AllocationID matches the barcode's order.
#
# READ-ONLY by design: the workbook is formula-driven (XLOOKUP/COUNTIF/stacked
# sheets). Writing from R would strip Excel's cached formula results. The app
# reads the values Excel last saved and re-reads automatically when the file
# changes (save in Excel -> app refreshes within a few seconds).
#
# Setup (once):
#   install.packages(c("shiny", "bslib", "readxl", "DT", "dplyr", "tidyr", "htmltools"))
# Run on your own computer (desktop mode):
#   shiny::runApp("path/to/folder/with/app.R")
#   Then click "File" (top right) and paste the workbook's full path. The app
#   remembers the last path that loaded successfully.
#
# Share it with the lab (server mode): edit and run run_server.R. In server mode
# the workbook paths are fixed by whoever starts the app, and the browser has no
# way to open, browse or see files: no File button, no raw Sheets tab, no file
# paths or raw error messages. Optional login: see LAB_TRACKER_USERS below.
# =============================================================================

suppressPackageStartupMessages({
  library(shiny)
  library(bslib)
  library(readxl)
  library(DT)
  library(dplyr)
  library(tidyr)
})

# Other packages loaded in the same R session (Bioconductor's S4Vectors /
# IRanges / AnnotationDbi, data.table, plyr, MASS ...) can mask these dplyr /
# tidyr functions and break the app ("unable to find an inherited method for
# function 'first'"). Pin them to dplyr / tidyr.
first     <- dplyr::first
last      <- dplyr::last
select    <- dplyr::select
filter    <- dplyr::filter
rename    <- dplyr::rename
slice     <- dplyr::slice
count     <- dplyr::count
desc      <- dplyr::desc
mutate    <- dplyr::mutate
summarise <- dplyr::summarise
arrange   <- dplyr::arrange
distinct  <- dplyr::distinct
coalesce  <- dplyr::coalesce
n         <- dplyr::n
separate_rows <- tidyr::separate_rows

POLL_MS <- 5000

# Hosted copy (Posit / shinyapps.io): settings live in lab_tracker_settings.R
# next to app.R, which sets the same environment variables run_server.R does.
if (file.exists("lab_tracker_settings.R")) source("lab_tracker_settings.R", local = TRUE)
# Lets the publisher see that the optional login package is needed.
if (FALSE) library(shinymanager)
if (FALSE) library(writexl)   # Submission prep: Excel download

env_flag <- function(name, default) {
  v <- tolower(trimws(Sys.getenv(name, "")))
  if (!nzchar(v)) default else v %in% c("1", "true", "yes", "on")
}

# Desktop (default) or server. Set by run_server.R.
SERVER_MODE  <- identical(tolower(trimws(Sys.getenv("LAB_TRACKER_MODE", "desktop"))), "server")
# Copy / CSV / Excel buttons on tables
ALLOW_EXPORT <- env_flag("LAB_TRACKER_ALLOW_EXPORT", TRUE)

# Big tables scroll inside a box that fits the browser window, so the sideways
# scrollbar is always on screen and the column headers stay put.
TABLE_HEIGHT <- "calc(100vh - 260px)"

# Re-extraction marks: green check = already re-extracted, orange dash = meets criteria
REEX_DONE <- '<span style="color:#1E6B3A;font-weight:700" title="Re-extracted">\u2713</span>'
REEX_DUE  <- '<span style="color:#C26A00;font-weight:700" title="Meets re-extraction criteria">\u2013</span>'
reex_mark <- function(state) {
  out <- rep(NA_character_, length(state))
  out[state %in% "Meets criteria"] <- REEX_DUE
  out[state %in% "Re-extracted"]   <- REEX_DONE
  out
}

# Excel TRUE/FALSE cells are shown as these marks (escaped so Windows R reads them fine)
CHECK_YES <- "\u2713"
CHECK_NO  <- "\u2717"

# Colors any cell that is exactly a check mark: green for TRUE, red for FALSE
CHECK_JS <- DT::JS(sprintf(
  "function(row){ $('td', row).each(function(){ var t = $(this).text();
     if (t === '%s') $(this).css({'color':'#1E6B3A','font-weight':'700','text-align':'center'});
     else if (t === '%s') $(this).css({'color':'#A12C2C','font-weight':'700','text-align':'center'}); }); }",
  CHECK_YES, CHECK_NO))

# Accepts what people actually paste: Windows "Copy as path" adds quotes,
# macOS Option-Cmd-C doesn't; ~ is expanded; backslashes are fine.
clean_path <- function(p) {
  if (is.null(p) || !length(p) || is.na(p[1])) return("")
  p <- trimws(p[1])
  p <- trimws(gsub("^[\"']+|[\"']+$", "", p))
  if (!nzchar(p)) return("")
  path.expand(p)
}

# Three workbooks. Only the lab tracker is required.
#   tracker - Lab_Tracker_2026.xlsx
#   backup  - BackupInventoryDatabase.xlsx (every sheet named as a year, e.g. "2025")
#   old     - OLD_FinalReports.xlsx ("CSV" sheet)
ENV_VARS <- c(tracker = "LAB_TRACKER_PATH", backup = "BACKUP_INVENTORY_PATH", old = "OLD_FINALREPORTS_PATH")

# Paths that last loaded successfully are remembered per user, outside the app folder.
CONFIG_DIR    <- tools::R_user_dir("lab_tracker", "config")
PATHS_FILE    <- file.path(CONFIG_DIR, "paths.rds")
OLD_PATH_FILE <- file.path(CONFIG_DIR, "last_path.txt")   # from earlier versions

read_saved_paths <- function() {
  p <- list(tracker = "", backup = "", old = "")
  saved <- tryCatch(readRDS(PATHS_FILE), error = function(e) NULL)
  if (is.list(saved)) {
    for (n in names(p)) if (is.character(saved[[n]]) && length(saved[[n]]) == 1) p[[n]] <- saved[[n]]
  }
  if (!nzchar(p$tracker) && file.exists(OLD_PATH_FILE)) {
    x <- tryCatch(readLines(OLD_PATH_FILE, n = 1, warn = FALSE), error = function(e) character())
    if (length(x)) p$tracker <- x
  }
  p
}

save_paths <- function(p) {
  try({
    dir.create(CONFIG_DIR, recursive = TRUE, showWarnings = FALSE)
    saveRDS(p, PATHS_FILE)
  }, silent = TRUE)
}

# Startup: environment variables win, then the last paths that loaded.
INITIAL_PATHS <- local({
  p <- if (SERVER_MODE) list(tracker = "", backup = "", old = "") else read_saved_paths()
  for (n in names(ENV_VARS)) {
    e <- clean_path(Sys.getenv(ENV_VARS[[n]], ""))
    if (nzchar(e)) p[[n]] <- e
  }
  if (!nzchar(p$tracker)) p$tracker <- "Lab_Tracker_2026.xlsx"
  p
})

# Optional login. LAB_TRACKER_USERS = path to a CSV with columns user,password
# (one row per person). Keep that file somewhere only you can open.
USERS_FILE <- clean_path(Sys.getenv("LAB_TRACKER_USERS", ""))
USE_LOGIN  <- nzchar(USERS_FILE)
CREDENTIALS <- NULL
if (USE_LOGIN) {
  if (!requireNamespace("shinymanager", quietly = TRUE)) {
    stop("A login was requested (LAB_TRACKER_USERS) but the shinymanager package isn't installed. ",
         "Run install.packages(\"shinymanager\").")
  }
  if (!file.exists(USERS_FILE)) stop("LAB_TRACKER_USERS file not found: ", USERS_FILE)
  CREDENTIALS <- utils::read.csv(USERS_FILE, stringsAsFactors = FALSE, colClasses = "character",
                                 strip.white = TRUE)
  if (!all(c("user", "password") %in% names(CREDENTIALS)) || !nrow(CREDENTIALS)) {
    stop("LAB_TRACKER_USERS must be a CSV with columns 'user' and 'password' and at least one row")
  }
  CREDENTIALS <- CREDENTIALS[nzchar(CREDENTIALS$user) & nzchar(CREDENTIALS$password), c("user", "password")]
}

# -----------------------------------------------------------------------------
# Small helpers
# -----------------------------------------------------------------------------

# IDs are compared case-insensitively with spaces/hyphens removed, so
# "20-2513463-43" matches "202513463-43" and "fa08209390" matches "FA08209390".
norm_id <- function(x) toupper(gsub("[[:space:]-]", "", x))

# SampleIDs are "20-" + 7 digits. Excel sometimes drops the leading zero of the
# 7 digits ("20-908705" for "20-0908705"); that zero is restored before comparing.
norm_sid <- function(x) {
  y <- norm_id(x)                      # always character, even when empty
  i <- !is.na(y) & grepl("^20[0-9]{6}$", y)
  y[i] <- paste0("200", substring(y[i], 3))
  y
}

# Filler values that are not real SampleIDs
PLACEHOLDER_IDS <- c("0", "NA", "N/A", "#N/A")

# A real barcode: present and not a filler like "0" / "NA"
# a barcode has a digit and no spaces, so notes typed into a Barcode cell
# ("Awaiting GSA is currently in flux") aren't read as tubes
valid_bc <- function(x) nz(x) & !norm_id(x) %in% PLACEHOLDER_IDS &
  grepl("[0-9]", coalesce(x, "")) & !grepl("\\s", trimws(coalesce(x, "")))

real_sid <- function(x) {
  x <- sub(",.*$", "", x)
  x[norm_id(x) %in% PLACEHOLDER_IDS] <- NA_character_
  x
}

nz <- function(x) !is.na(x) & x != ""

collapse_u <- function(x, sep = ", ") {
  x <- unique(x[nz(x)])
  if (length(x)) paste(x, collapse = sep) else NA_character_
}

col_or_na <- function(df, name) {
  if (name %in% names(df)) df[[name]] else rep(NA_character_, nrow(df))
}

# First column whose name matches any candidate, ignoring case, spaces and "/"
# (so a renamed header like "GSA (Result)" vs "GSA/WGS(Result)" still works).
find_col <- function(df, candidates) {
  key <- function(x) tolower(gsub("[[:space:]/]", "", x))
  hit <- match(key(candidates), key(names(df)))
  hit <- hit[!is.na(hit)]
  if (length(hit)) names(df)[hit[1]] else NA_character_
}

col_any <- function(df, candidates) {
  nm <- find_col(df, candidates)
  if (is.na(nm)) rep(NA_character_, nrow(df)) else df[[nm]]
}

# Allocations that went straight to USUHS / NCRAD / Columbia University (CU,
# biomarker): the tube was never received in the lab and is gone for good.
# Recognised from Original List Status ("... Sent to USUHS") or Comments
# ("USUHS_Only", "BIOMARKER_ONLY", "NCRAD").
sent_dest <- function(status, comments) {
  s <- tolower(coalesce(status, "")); c <- tolower(coalesce(comments, ""))
  case_when(grepl("sent to usuhs", s)    | grepl("usuhs[_ ]*only", c)     ~ "USUHS",
            grepl("sent to ncrad", s)    | grepl("^\\s*ncrad", c)        ~ "NCRAD",
            grepl("sent to columbia", s) | grepl("biomarker[_ ]*only", c) ~ "CU",
            TRUE ~ NA_character_)
}
# per SampleID + allocation: where that allocation's tube went (USUHS / NCRAD / CU)
sent_map <- function(ol) {
  if (is.null(ol) || !nrow(ol)) return(tibble(sid_key = character(), a = character(), dest = character()))
  ol %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
              dest = sent_dest(col_or_na(., "Status"), col_or_na(., "Comments"))) %>%
    filter(nz(a), !is.na(dest)) %>%
    distinct(sid_key, a, .keep_all = TRUE)
}
NEVER_RECEIVED <- function(dest) paste(dest, "only (never received)")

# "Same SampleID, other allocation" for a row with NO barcode (e.g. a sample
# dropped in this allocation): the sample's tubes in other allocations, then its
# other Original List allocations no tube covers, with their Status
# ("38754 (Dropped: Low Amount)"). Same wording as for barcode rows.
other_alloc_label <- function(sid_keys, this_allocs, bt, ol) {
  out <- rep(NA_character_, length(sid_keys))
  if (!length(sid_keys)) return(out)
  q <- tibble(i = seq_along(sid_keys), sid_key = sid_keys, this = alloc_base(this_allocs))
  tubes <- bt %>% filter(nz(SampleID), valid_bc(Barcode)) %>%
    transmute(sid_key = norm_sid(SampleID), Barcode, Allocation,
              a = alloc_base(sub(",.*$", "", coalesce(Allocation, ""))),
              Sent_To = if ("Sent_To" %in% names(bt)) Sent_To else NA_character_)
  olr <- if (!is.null(ol) && nrow(ol)) ol %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
              st = col_or_na(., "Status")) %>% filter(nz(a))
  else tibble(sid_key = character(), a = character(), st = character())
  tb <- q %>% inner_join(tubes, by = "sid_key", relationship = "many-to-many") %>%
    filter(coalesce(a, "") != coalesce(this, "")) %>%
    mutate(lbl = paste0(Barcode, " (", coalesce(Allocation, "?"),
                        ifelse(is.na(Sent_To), "", paste0(", ", NEVER_RECEIVED(Sent_To))), ")")) %>%
    group_by(i) %>% summarise(t = collapse_u(lbl), .groups = "drop")
  oo <- q %>% inner_join(olr, by = "sid_key", relationship = "many-to-many") %>%
    filter(a != coalesce(this, "")) %>%
    anti_join(tubes %>% distinct(sid_key, a), by = c("sid_key", "a")) %>%
    group_by(i, a) %>% summarise(st = collapse_u(st, "; "), .groups = "drop") %>%
    mutate(lbl = ifelse(is.na(st), a, paste0(a, " (", st, ")"))) %>%
    group_by(i) %>% summarise(o = paste(lbl, collapse = ", "), .groups = "drop")
  x <- q %>% left_join(tb, by = "i") %>% left_join(oo, by = "i")
  # as.character(): when nothing else is found both parts are all-NA logical, and a
  # logical column breaks bind_rows() with the character tube rows
  as.character(ifelse(is.na(x$t), x$o, ifelse(is.na(x$o), x$t, paste(x$t, x$o, sep = ", "))))
}

# Names that appear in allocation columns but are rerun lists, not allocations
NON_ALLOCATION <- function(x) grepl("re-?do|rerun", x, ignore.case = TRUE)

# Drop those from a single value or a comma-separated list; NA if nothing is left
clean_alloc <- function(x) {
  vapply(strsplit(coalesce(x, ""), ",\\s*"), function(a) {
    a <- trimws(a); collapse_u(a[nz(a) & !NON_ALLOCATION(a)])
  }, character(1))
}

# Sheets tab: only these main sheets (in this order, if present) ...
MAIN_SHEETS <- c("Original List", "ALL_FinalReports", "GSA", "WGS", "Awaiting_GSA",
                 "Sanger_Tracker", "Overview_Tracker_2.0")
# ... plus, under its own "Allocations" sub-tab, the one-sheet-per-allocation
# sheets: names that are an allocation number, e.g. "38897" or "38329_1".
is_alloc_sheet <- function(x) grepl("^[0-9]{4,}(_[0-9]+)?$", trimws(x))

# Original List: GSA result column, under its current or earlier name. It's an
# educated guess, not a confirmed result, so the app never reads or shows it.
GSA_RESULT_NAMES <- c("GSA (Result)", "GSA(Result)", "GSA/WGS(Result)")
drop_gsa_result <- function(df) {
  key <- function(x) tolower(gsub("[^A-Za-z0-9]", "", x))
  if (is.null(df) || !ncol(df)) return(df)
  df[, !key(names(df)) %in% key(GSA_RESULT_NAMES), drop = FALSE]
}

require_cols <- function(df, cols, sheet) {
  miss <- setdiff(cols, names(df))
  if (length(miss)) {
    stop(sprintf("Sheet '%s' is missing column(s): %s", sheet, paste(miss, collapse = ", ")))
  }
  invisible(df)
}

# One Excel cell (from col_types = "list") -> character. Dates become YYYY-MM-DD.
cell_chr <- function(v) {
  if (is.null(v) || length(v) == 0) return(NA_character_)
  if (length(v) == 1 && is.na(v)) return(NA_character_)
  if (inherits(v, "POSIXt")) return(format(v, "%Y-%m-%d", tz = "UTC"))
  if (is.numeric(v)) return(format(v, scientific = FALSE, trim = TRUE, digits = 15))
  if (is.logical(v)) return(if (isTRUE(v)) CHECK_YES else CHECK_NO)  # Excel checkboxes
  # [\h\v] also strips non-breaking spaces, so a cell holding only one counts as empty
  s <- trimws(as.character(v), whitespace = "[\\h\\v]")
  if (identical(s, "")) NA_character_ else s
}

# "38329_1", "38329_2" ... are part of allocation 38329. Only a trailing
# underscore + digits is stripped, so names like "OLD_DUKE" are left alone.
alloc_base <- function(x) sub("_[0-9]+$", "", trimws(x))

alloc_base_list <- function(x) {
  vapply(strsplit(coalesce(x, ""), ",\\s*"),
         function(s) collapse_u(alloc_base(s)), character(1))
}

ffill <- function(x) {
  if (length(x) > 1) for (i in 2:length(x)) if (is.na(x[i])) x[i] <- x[i - 1]
  x
}

# Whole sheet as an all-character data frame (no header handling yet).
# Reading every cell as "list" keeps mixed columns intact, e.g. GSA "Handed Back"
# holds both "Yes" and dates; a typed read would silently turn one of them to NA.
read_raw <- function(path, sheet) {
  raw <- read_excel(path, sheet = sheet, col_names = FALSE, col_types = "list",
                    .name_repair = "minimal")
  if (!ncol(raw)) return(data.frame())
  out <- lapply(raw, function(col) vapply(col, cell_chr, character(1), USE.NAMES = FALSE))
  out <- as.data.frame(out, stringsAsFactors = FALSE, optional = TRUE)
  names(out) <- paste0("X", seq_along(out))
  out
}

find_header <- function(m, match, sheet) {
  hr <- which(m[[1]] == match)[1]
  if (is.na(hr)) stop(sprintf("Sheet '%s': couldn't find header '%s' in column A", sheet, match))
  hr
}

build_table <- function(m, names_vec, first_data, stop_at_blank = TRUE) {
  names_vec[!nz(names_vec)] <- NA
  keep <- seq_along(names_vec)
  if (stop_at_blank) {
    b <- which(is.na(names_vec))
    if (length(b)) keep <- seq_len(b[1] - 1)
  }
  d <- if (first_data > nrow(m)) m[0, keep, drop = FALSE] else m[first_data:nrow(m), keep, drop = FALSE]
  nm <- names_vec[keep]
  unnamed <- is.na(nm)
  nm[unnamed] <- paste0("col", keep[unnamed])
  names(d) <- make.unique(nm, sep = "_")
  # drop unnamed columns that are entirely empty, then entirely empty rows
  empty_unnamed <- unnamed & vapply(d, function(x) all(is.na(x)), logical(1))
  d <- d[, !empty_unnamed, drop = FALSE]
  d <- d[rowSums(!is.na(d)) > 0, , drop = FALSE]
  rownames(d) <- NULL
  d
}

# Single header row; columns stop at the first blank header (drops side tables
# like Original List's "Allocations:" helper block and the "Find Samples" cells).
read_simple <- function(path, sheet, header_match, stop_at_blank = TRUE) {
  m <- read_raw(path, sheet)
  hr <- find_header(m, header_match, sheet)
  build_table(m, unlist(m[hr, ], use.names = FALSE), hr + 1, stop_at_blank)
}

# Overview_Tracker_2.0: section row (Sanger / GSA / WGS) above the header row.
read_overview <- function(path, sheet = "Overview_Tracker_2.0") {
  m <- read_raw(path, sheet)
  hr <- find_header(m, "Allocation", sheet)
  hdr <- unlist(m[hr, ], use.names = FALSE)
  sec <- if (hr > 1) ffill(unlist(m[hr - 1, ], use.names = FALSE)) else rep(NA_character_, length(hdr))
  nm <- ifelse(is.na(sec) | is.na(hdr), hdr, paste(sec, hdr, sep = " · "))
  build_table(m, nm, hr + 1, stop_at_blank = TRUE)
}

# Sanger_Tracker: stage row with an APOE / ABCA7 sub-row beneath it.
read_sanger <- function(path, sheet = "Sanger_Tracker") {
  m <- read_raw(path, sheet)
  hr <- find_header(m, "Allocation", sheet)
  top <- unlist(m[hr, ], use.names = FALSE)
  sub <- unlist(m[hr + 1, ], use.names = FALSE)
  nm <- ifelse(sub %in% c("APOE", "ABCA7"), paste(ffill(top), sub), ifelse(is.na(top), sub, top))
  build_table(m, nm, hr + 2, stop_at_blank = FALSE)
}

# Fast reader for large, date-free sheets: every cell read as text.
read_text_table <- function(path, sheet, header_match) {
  raw <- read_excel(path, sheet = sheet, col_names = FALSE, col_types = "text",
                    .name_repair = "minimal")
  if (!ncol(raw)) stop(sprintf("Sheet '%s' is empty", sheet))
  m <- as.data.frame(lapply(raw, function(x) { x <- trimws(x, whitespace = "[\\h\\v]"); x[x == ""] <- NA_character_; x }),
                     stringsAsFactors = FALSE, optional = TRUE)
  names(m) <- paste0("X", seq_along(m))
  hr <- find_header(m, header_match, sheet)
  build_table(m, unlist(m[hr, ], use.names = FALSE), hr + 1, stop_at_blank = TRUE)
}

# "Requested Reruns" table (optional): samples someone asked to have rerun (e.g.
# WGS discrepancies). It sits in the Sanger_Rerun_List tab, to the right of the
# log with a blank column between, under a title cell "Requested Reruns"; the
# header row below the title has SampleID and Gene (APOE, ABCA7, Both,
# "APOE/ABCA7"...), optionally Allocation, Requester and Note. A separate
# "Requested Reruns" sheet is read the same way.
# Returns one row per sample + gene, plus a Problem column. A Result of Pass or
# Done means the request is finished: the sample leaves the Sanger Hiatus list.
REQ_TITLE <- "Requested Reruns"
# a request is finished when its Result says Pass or Done
req_closed <- function(result) {
  r <- trimws(coalesce(result, ""))
  coalesce(sanger_cat(r) == "Pass", FALSE) | grepl("^done\\b", r, ignore.case = TRUE)
}
read_requested <- function(path, sheets) {
  empty <- tibble(sid_key = character(), SampleID = character(), g = character(),
                  Allocation = character(), Requester = character(), Note = character(),
                  Result = character(), Resubmitted = character(),
                  Input = character(), GeneText = character(), Problem = character())
  if (is.null(path) || !length(sheets)) return(empty)
  key <- function(x) gsub("[^a-z0-9]", "", tolower(coalesce(x, "")))
  sh_rl  <- sheets[key(sheets) == key("Sanger_Rerun_List")]
  sh_req <- sheets[key(sheets) == key(REQ_TITLE)]
  m <- NULL; tr <- NA; tc <- NA
  for (x in c(sh_rl, sh_req)) {
    mm <- read_raw(path, x)
    if (!nrow(mm)) next
    km <- matrix(key(as.matrix(mm)), nrow = nrow(mm))
    hit <- which(km == key(REQ_TITLE), arr.ind = TRUE)
    if (!nrow(hit) && x %in% sh_req) hit <- cbind(0L, 1L)   # own sheet, no title: start at A
    if (nrow(hit)) { m <- mm; tr <- hit[1, 1]; tc <- hit[1, 2]; break }
  }
  if (is.null(m)) return(empty)
  # header row: first row below the title with SampleID in the title's column
  # (or the column next to it)
  hr <- NA; hc <- NA
  for (r in seq(tr + 1, min(nrow(m), tr + 4))) {
    for (cc in unique(pmin(ncol(m), c(tc, tc + 1)))) {
      if (key(m[[cc]][r]) %in% c("sampleid", "sample")) { hr <- r; hc <- cc; break }
    }
    if (!is.na(hr)) break
  }
  if (is.na(hr)) stop("Requested Reruns: found the title but no 'SampleID' header under it")
  m <- m[, hc:ncol(m), drop = FALSE]
  d <- build_table(m, unlist(m[hr, ], use.names = FALSE), hr + 1, stop_at_blank = TRUE)
  d <- tibble(Input = col_any(d, c("SampleID", "Sample ID", "Sample_ID", "Sample")),
              GeneText = col_any(d, c("Gene", "Genes")),
              Allocation = col_any(d, c("Allocation", "AllocationID")),
              Requester = col_any(d, c("Requester", "Requested by", "Requestor")),
              Note = col_any(d, c("Note", "Notes", "Comment", "Comments")),
              Result = col_any(d, c("Result", "Results", "Rerun result")),
              Resubmitted = col_any(d, c("RE-Submission Date", "Resubmission Date", "Submission Date", "Date"))) %>%
    filter(nz(Input))
  if (!nrow(d)) return(empty)
  gt <- toupper(coalesce(d$GeneText, ""))
  a  <- grepl("APOE", gt); b <- grepl("ABCA7", gt); both <- grepl("BOTH", gt)
  blank <- !nzchar(trimws(gt))
  d$Problem <- case_when(blank ~ "Gene is blank (listed for both)",
                         !(a | b | both) ~ "Gene not recognised (listed for both)",
                         TRUE ~ NA_character_)
  d$sid_key <- norm_sid(d$Input)
  bad_id <- !grepl("^20[0-9]{7}$", d$sid_key)
  id_msg <- "SampleID doesn't look like 20-xxxxxxx"
  d$Problem <- case_when(bad_id & !is.na(d$Problem) ~ paste0(d$Problem, "; ", id_msg),
                         bad_id                     ~ id_msg,
                         TRUE                       ~ d$Problem)
  use_a <- a | both | !(a | b)
  use_b <- b | both | !(a | b)
  bind_rows(d[use_a, ] %>% mutate(g = "APOE"), d[use_b, ] %>% mutate(g = "ABCA7")) %>%
    mutate(SampleID = if_else(grepl("^20[0-9]{7}$", sid_key),
                              paste0(substr(sid_key, 1, 2), "-", substr(sid_key, 3, 9)), Input)) %>%
    distinct(sid_key, g, .keep_all = TRUE) %>%
    select(sid_key, SampleID, g, Allocation, Requester, Note, Result, Resubmitted, Input, GeneText, Problem)
}

# "Re-extracted Samples" table (optional) in the Sanger_Rerun_List tab: a title
# cell "Re-extracted Samples" with the SampleIDs under it and their new
# allocation in the next column (header "AllocationID" next to the title).
RX_TITLE <- "Re-extracted Samples"
read_reextracted <- function(path, sheets) {
  empty <- tibble(sid_key = character(), SampleID = character(), Allocation = character())
  if (is.null(path) || !length(sheets)) return(empty)
  key <- function(x) gsub("[^a-z0-9]", "", tolower(coalesce(x, "")))
  sh <- sheets[key(sheets) == key("Sanger_Rerun_List")]
  if (!length(sh)) return(empty)
  m <- read_raw(path, sh[1])
  if (!nrow(m)) return(empty)
  km <- matrix(key(as.matrix(m)), nrow = nrow(m))
  hit <- which(km == key(RX_TITLE), arr.ind = TRUE)
  if (!nrow(hit)) return(empty)
  tr <- hit[1, 1]; tc <- hit[1, 2]
  if (tr >= nrow(m)) return(empty)
  ids <- m[[tc]][(tr + 1):nrow(m)]
  al  <- if (tc < ncol(m)) m[[tc + 1]][(tr + 1):nrow(m)] else rep(NA_character_, length(ids))
  d <- tibble(SampleID = trimws(ids), Allocation = trimws(al)) %>%
    filter(nz(SampleID), !key(SampleID) %in% c("sampleid", "sample")) %>%
    mutate(sid_key = norm_sid(SampleID),
           Allocation = if_else(key(Allocation) %in% c("allocationid", "allocation"), NA_character_, Allocation))
  d %>% filter(grepl("^[0-9]", sid_key)) %>% distinct(sid_key, .keep_all = TRUE) %>%
    select(sid_key, SampleID, Allocation)
}

# BackupInventoryDatabase: one sheet per year ("2025", later "2026", ...).
read_backup <- function(path) {
  yrs <- grep("^[0-9]{4}$", excel_sheets(path), value = TRUE)
  if (!length(yrs)) stop("No year sheet (e.g. '2025') found in the backup inventory workbook")
  out <- bind_rows(lapply(yrs, function(sh) {
    d <- read_text_table(path, sh, "Allocation")
    d$Inventory_Sheet <- sh
    d
  }))
  require_cols(out, c("Barcode", "Plate"), "Backup inventory")
}

# OLD_FinalReports: the stacked "CSV" sheet of older biorepository final reports.
read_old_fr <- function(path) {
  out <- read_text_table(path, "CSV", "order_id")
  require_cols(out, c("order_id", "SampleID", "barcode_2d"), "OLD_FinalReports CSV")
}

# Optional sheets shouldn't take the whole app down if renamed.
try_read <- function(expr) {
  tryCatch(expr, error = function(e) structure(data.frame(), load_error = conditionMessage(e)))
}

# -----------------------------------------------------------------------------
# Result categories
# -----------------------------------------------------------------------------

# Pass/PASS and Fail/FAIL are collapsed for display. Any result containing
# "dropped" or "no repeat" ("Fail, Dropped", "Fail, No Repeat", or the older
# hyphenated "F - ail, dropped" / "F-ail")
# means the sample failed but was dropped, so it won't be repeated. It is shown
# as "Fail, dropped" and not counted among fails that still need a redo.
sanger_cat <- function(x) {
  y <- trimws(x)
  case_when(
    !nz(y)                                   ~ NA_character_,
    grepl("|", y, fixed = TRUE)              ~ "Multiple",
    grepl("^pass", y, ignore.case = TRUE)    ~ "Pass",
    grepl("dropped|no\\s*repeat", y, ignore.case = TRUE) ~ "Fail, dropped",   # "Fail, Dropped", "Fail, No Repeat"
    grepl("^F\\s*-\\s*ail", y, ignore.case = TRUE) ~ "Fail, dropped",   # older "F - ail" form
    grepl("^fail", y, ignore.case = TRUE)    ~ "Fail",
    toupper(y) %in% c("N/A", "NA")           ~ "N/A",
    TRUE                                     ~ y
  )
}

is_date_str <- function(x) nz(x) & grepl("^\\d{4}-\\d{2}-\\d{2}$|^\\d{1,2}/\\d{1,2}/\\d{2,4}$", trimws(x))
# Rerun dates as text: "2026-08-01" or "8/1/2026" -> "08/01/2026"; several
# joined in date order: "04/01/2026, 08/01/2026".
as_day <- function(x) {
  x <- trimws(as.character(x))
  d <- suppressWarnings(as.Date(x, format = "%Y-%m-%d"))
  d2 <- suppressWarnings(as.Date(x, format = "%m/%d/%Y"))
  d[is.na(d)] <- d2[is.na(d)]
  d
}
date_list <- function(x) {
  d <- sort(unique(as_day(x)))
  if (!length(d)) NA_character_ else paste(format(d, "%m/%d/%Y"), collapse = ", ")
}

is_returned <- function(x) nz(x) & (toupper(x) == "YES" | is_date_str(x))

TONES <- c(
  "Pass" = "ok", "Returned" = "ok",
  "Fail" = "bad", "Mismatch" = "bad",
  "Fail, dropped" = "drop", "Multiple" = "flag",
  "At GSA" = "warn", "At WGS" = "warn",
  "Awaiting GSA" = "info", "Sent to USUHS" = "info",
  "N/A" = "muted", "Not sent" = "muted",
  "In backup" = "ok", "Removed" = "flag"
)

pill <- function(x, title = NULL) {
  tone <- unname(TONES[x])
  tone[is.na(tone)] <- "muted"
  ttl <- if (is.null(title)) "" else ifelse(nz(title), sprintf(' title="%s"', htmltools::htmlEscape(title, attribute = TRUE)), "")
  out <- sprintf('<span class="pill pill-%s"%s>%s</span>', tone, ttl, htmltools::htmlEscape(x))
  out[is.na(x)] <- ""
  out
}

# Colored pill that keeps the sheet's own wording ("Fail (Sex Mismatch)" stays
# readable) but takes its color from the result category: Pass green, Fail red,
# "Fail, dropped" muted red, anything else grey (sheet text on hover).
result_pill <- function(x) {
  cat  <- sanger_cat(x)
  tone <- unname(TONES[cat])
  tone[is.na(tone)] <- "muted"
  label <- ifelse(cat %in% "Fail, dropped", "Fail, dropped", coalesce(x, ""))
  out <- sprintf('<span class="pill pill-%s" title="%s">%s</span>', tone,
                 htmltools::htmlEscape(coalesce(x, ""), attribute = TRUE), htmltools::htmlEscape(label))
  out[!nz(x)] <- ""
  out
}

# -----------------------------------------------------------------------------
# The barcode table
# -----------------------------------------------------------------------------

build_barcode_table <- function(afr, ol, gsa, wgs, aw, bk = NULL, old = NULL) {
  require_cols(afr, c("order_id", "SampleID", "child_aliquot", "barcode_2d"), "ALL_FinalReports")
  require_cols(ol, c("Sample_ID", "Status", "AllocationID"), "Original List")
  require_cols(gsa, c("Handoff DATE", "Barcode", "SampleID", "Allocation", "Handed Back"), "GSA")
  require_cols(wgs, c("Handoff DATE", "Barcode", "SampleID", "Allocation", "Handed Back"), "WGS")
  
  afr_b <- afr %>%
    filter(valid_bc(barcode_2d)) %>%
    mutate(key = norm_id(barcode_2d),
           plate_id = col_or_na(., "plate_id"), position = col_or_na(., "position")) %>%
    group_by(key) %>%
    summarise(
      afr_barcode = first(barcode_2d),
      afr_sid     = collapse_u(SampleID),
      afr_alloc   = collapse_u(order_id),
      afr_aliquot = collapse_u(child_aliquot),
      Plate       = collapse_u(plate_id),
      Position    = collapse_u(position),
      .groups = "drop"
    )
  
  gsa_b <- gsa %>%
    filter(valid_bc(Barcode)) %>%
    mutate(key = norm_id(Barcode), AliquotID = col_or_na(., "AliquotID")) %>%
    group_by(key) %>%
    summarise(
      gsa_barcode = first(Barcode),
      gsa_sid     = collapse_u(SampleID),
      gsa_alloc   = collapse_u(Allocation),
      gsa_aliquot = collapse_u(AliquotID),
      GSA_Handoff = collapse_u(`Handoff DATE`),
      GSA_Back    = collapse_u(`Handed Back`),
      .groups = "drop"
    )
  
  wgs_b <- wgs %>%
    filter(valid_bc(Barcode)) %>%
    mutate(key = norm_id(Barcode), AliquotID = col_or_na(., "AliquotID"),
           Type = col_or_na(., "Type"), Notes = col_or_na(., "Notes")) %>%
    group_by(key) %>%
    summarise(
      wgs_barcode = first(Barcode),
      wgs_sid     = collapse_u(SampleID),
      wgs_alloc   = collapse_u(Allocation),
      wgs_aliquot = collapse_u(AliquotID),
      WGS_Handoff = collapse_u(`Handoff DATE`),
      WGS_Type    = collapse_u(Type),
      WGS_Back    = collapse_u(`Handed Back`),
      WGS_Notes   = collapse_u(Notes),
      .groups = "drop"
    )
  
  aw_b <- if (nrow(aw) && "Barcode" %in% names(aw)) {
    aw %>%
      filter(valid_bc(Barcode)) %>%
      mutate(key = norm_id(Barcode),
             SampleID = col_or_na(., "SampleID"), AliquotID = col_or_na(., "AliquotID"),
             Allocation = col_or_na(., "Allocation"), DateAdded = col_or_na(., "Date Added"),
             Pos = col_or_na(., "TEMP_Backup Inventory Plate POS")) %>%
      group_by(key) %>%
      summarise(
        aw_barcode = first(Barcode), aw_sid = collapse_u(SampleID),
        aw_aliquot = collapse_u(AliquotID), aw_alloc = collapse_u(Allocation),
        Awaiting_Added = collapse_u(DateAdded), Awaiting_Pos = collapse_u(Pos),
        .groups = "drop"
      )
  } else {
    tibble(key = character(), aw_barcode = character(), aw_sid = character(),
           aw_aliquot = character(), aw_alloc = character(),
           Awaiting_Added = character(), Awaiting_Pos = character())
  }
  
  old_b <- if (!is.null(old) && nrow(old)) {
    old %>%
      filter(valid_bc(barcode_2d)) %>%
      mutate(key = norm_id(barcode_2d),
             child_aliquot = col_or_na(., "child_aliquot"), SourceFile = col_or_na(., "SourceFile")) %>%
      group_by(key) %>%
      summarise(
        old_barcode = first(barcode_2d),
        old_sid     = collapse_u(SampleID),
        old_alloc   = collapse_u(order_id),
        old_aliquot = collapse_u(child_aliquot),
        Old_Report  = collapse_u(SourceFile),
        .groups = "drop"
      )
  } else {
    tibble(key = character(), old_barcode = character(), old_sid = character(),
           old_alloc = character(), old_aliquot = character(), Old_Report = character())
  }
  
  bk_b <- if (!is.null(bk) && nrow(bk)) {
    bk %>%
      filter(valid_bc(Barcode)) %>%
      mutate(key = norm_id(Barcode),
             SampleID = col_or_na(., "SampleID"), AliquotID = col_or_na(., "AliquotID"),
             Allocation = col_or_na(., "Allocation"), Note = col_or_na(., "Note"),
             Pos = col_or_na(., "Backup Inventory Plate POS")) %>%
      group_by(key) %>%
      summarise(
        bk_barcode   = first(Barcode),
        bk_sid       = collapse_u(SampleID),
        bk_alloc     = collapse_u(Allocation),
        bk_aliquot   = collapse_u(AliquotID),
        Backup_Plate = collapse_u(Plate),
        Backup_Pos   = collapse_u(Pos),
        Backup_Note  = collapse_u(Note, " | "),
        .groups = "drop"
      )
  } else {
    tibble(key = character(), bk_barcode = character(), bk_sid = character(),
           bk_alloc = character(), bk_aliquot = character(), Backup_Plate = character(),
           Backup_Pos = character(), Backup_Note = character())
  }
  
  bt <- afr_b %>%
    full_join(old_b, by = "key") %>%
    full_join(gsa_b, by = "key") %>%
    full_join(wgs_b, by = "key") %>%
    full_join(aw_b,  by = "key") %>%
    full_join(bk_b,  by = "key") %>%
    mutate(
      Barcode    = coalesce(afr_barcode, old_barcode, gsa_barcode, wgs_barcode, aw_barcode, bk_barcode),
      in_afr     = !is.na(afr_barcode),
      in_old     = !is.na(old_barcode),
      in_gsa     = !is.na(gsa_barcode),
      in_wgs     = !is.na(wgs_barcode),
      in_aw      = !is.na(aw_barcode),
      in_backup  = !is.na(bk_barcode),
      Allocation_Raw = coalesce(clean_alloc(afr_alloc), clean_alloc(old_alloc), clean_alloc(gsa_alloc),
                                clean_alloc(wgs_alloc), clean_alloc(aw_alloc), clean_alloc(bk_alloc)),
      Allocation = alloc_base_list(Allocation_Raw),
      Aliquot    = coalesce(afr_aliquot, old_aliquot, gsa_aliquot, wgs_aliquot, aw_aliquot, bk_aliquot),
      # SampleID used to reach Original List: final reports first, then the logs,
      # then the backup inventory. Placeholders like "0" are skipped.
      SampleID   = coalesce(real_sid(afr_sid), real_sid(old_sid), real_sid(gsa_sid),
                            real_sid(wgs_sid), real_sid(aw_sid), real_sid(bk_sid)),
      Backup_Stage = case_when(
        in_backup & grepl("^\\s*removed", coalesce(Backup_Note, ""), ignore.case = TRUE) ~ "Removed",
        in_backup ~ "In backup",
        TRUE      ~ NA_character_
      )
    )
  
  # Sources that disagree about which SampleID a barcode belongs to.
  # Placeholders ("0") are ignored and a dropped leading zero is not a conflict.
  bt$ID_conflict <- as.character(unlist(mapply(function(a, o, g, w, x, b) {
    src <- c(FinalReports = a, `Old final reports` = o, GSA = g, WGS = w, Awaiting = x, Backup = b)
    src <- src[nz(src)]
    src <- src[!norm_id(src) %in% PLACEHOLDER_IDS]
    if (length(unique(norm_sid(src))) <= 1) return(NA_character_)
    paste(paste(names(src), src, sep = ": "), collapse = " / ")
  }, bt$afr_sid, bt$old_sid, bt$gsa_sid, bt$wgs_sid, bt$aw_sid, bt$bk_sid,
  USE.NAMES = FALSE, SIMPLIFY = FALSE)))
  
  # ---- bridge to Original List ----
  olk <- ol %>%
    filter(nz(Sample_ID)) %>%
    mutate(sid_key = norm_sid(Sample_ID),
           APOE_raw  = col_or_na(., "APOESanger(Result)"),
           ABCA7_raw = col_or_na(., "ABCA7Sanger(Result)"),
           Substudy  = col_or_na(., "Substudy"),
           Sex       = col_or_na(., "Sex"),
           Race      = col_or_na(., "Race"),
           Ethnicity = col_or_na(., "Ethnicity"),
           OL_Comments = col_or_na(., "Comments"))
  
  ol_link <- bt %>%
    select(key, SampleID, Allocation) %>%
    filter(nz(SampleID)) %>%
    mutate(sid_key = norm_sid(SampleID)) %>%
    inner_join(olk, by = "sid_key", relationship = "many-to-many")
  
  # Prefer the Original List row for this barcode's own allocation; a reused
  # SampleID in another allocation is a different request, not this tube.
  ol_link$alloc_match <- as.logical(unlist(mapply(function(a, set) {
    nz(a) && nz(set) && alloc_base(a) %in% alloc_base(strsplit(set, ",\\s*")[[1]])
  }, ol_link$AllocationID, ol_link$Allocation, USE.NAMES = FALSE, SIMPLIFY = FALSE)))
  
  # Status and results come ONLY from the Original List row for this tube's own
  # allocation. If that allocation isn't in Original List (e.g. an older
  # allocation), they stay blank rather than borrowing another allocation's row.
  # Sample-level fields (substudy, sex, race, ethnicity) may come from any row.
  ol_b <- ol_link %>%
    group_by(key) %>%
    summarise(
      OL_match      = if (any(alloc_match)) "SampleID + allocation" else "Allocation not in Original List",
      Status        = collapse_u(Status[alloc_match], " | "),
      OL_Allocation = collapse_u(AllocationID[alloc_match], " | "),
      APOE_raw      = collapse_u(APOE_raw[alloc_match], " | "),
      ABCA7_raw     = collapse_u(ABCA7_raw[alloc_match], " | "),
      OL_Comments   = collapse_u(OL_Comments[alloc_match], " | "),
      Substudy      = collapse_u(Substudy[if (any(alloc_match)) alloc_match else TRUE], " | "),
      Sex           = collapse_u(Sex[if (any(alloc_match)) alloc_match else TRUE], " | "),
      Race          = collapse_u(Race[if (any(alloc_match)) alloc_match else TRUE], " | "),
      Ethnicity     = collapse_u(Ethnicity[if (any(alloc_match)) alloc_match else TRUE], " | "),
      .groups = "drop"
    )
  
  # ---- other barcodes carrying the same SampleID ----
  # Same allocation      -> sample was re-tubed (new barcode)
  # Different allocation -> sample reallocated / re-requested
  # Allocation suffixes ("38329_1") count as the parent allocation ("38329");
  # Allocation is already the parent ID here.
  # Placeholder SampleIDs like "0" are ignored.
  sm <- sent_map(ol)
  bt <- bt %>%
    mutate(sid_key_tmp = norm_sid(coalesce(SampleID, "")),
           a_tmp = alloc_base(sub(",.*$", "", coalesce(Allocation, "")))) %>%
    left_join(sm %>% rename(sid_key_tmp = sid_key, a_tmp = a, Sent_To = dest),
              by = c("sid_key_tmp", "a_tmp")) %>%
    select(-sid_key_tmp, -a_tmp)
  sib <- bt %>%
    filter(nz(SampleID), !norm_id(SampleID) %in% PLACEHOLDER_IDS) %>%
    transmute(sid_key = norm_sid(SampleID), key, Barcode, Allocation, Sent_To)
  pairs <- sib %>%
    inner_join(sib, by = "sid_key", suffix = c("", ".o"), relationship = "many-to-many") %>%
    filter(key != key.o)
  pairs$same_alloc <- as.logical(unlist(mapply(function(a, b) {
    if (!nz(a) || !nz(b)) return(FALSE)
    length(intersect(alloc_base(strsplit(a, ",\\s*")[[1]]),
                     alloc_base(strsplit(b, ",\\s*")[[1]]))) > 0
  }, pairs$Allocation, pairs$Allocation.o, USE.NAMES = FALSE, SIMPLIFY = FALSE)))
  sib_b <- pairs %>%
    group_by(key) %>%
    summarise(
      Retubed_With   = collapse_u(Barcode.o[same_alloc]),
      Other_Alloc_Barcodes = collapse_u(paste0(Barcode.o, " (", coalesce(Allocation.o, "?"),
                                               ifelse(is.na(Sent_To.o), "", paste0(", ", NEVER_RECEIVED(Sent_To.o))),
                                               ")")[!same_alloc]),
      .groups = "drop"
    )
  
  # Original List allocations of the same SampleID that no barcode covers yet
  # (e.g. reallocated, final report not in) -> "37677 (Dropped: Low Amount)"
  sid_allocs <- sib %>%
    separate_rows(Allocation, sep = ",\\s*") %>%
    filter(nz(Allocation)) %>%
    distinct(sid_key, a = Allocation)
  ol_only <- ol %>%
    filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
              st = col_or_na(., "Status")) %>%
    filter(nz(a)) %>%
    anti_join(sid_allocs, by = c("sid_key", "a")) %>%
    group_by(sid_key, a) %>%
    summarise(st = collapse_u(st, "; "), .groups = "drop") %>%
    mutate(label = ifelse(is.na(st), a, paste0(a, " (", st, ")"))) %>%
    group_by(sid_key) %>%
    summarise(OL_only = paste(label, collapse = ", "), .groups = "drop")
  ol_only_b <- sib %>% select(key, sid_key) %>% inner_join(ol_only, by = "sid_key") %>% select(key, OL_only)
  
  bt %>%
    left_join(ol_b, by = "key") %>%
    left_join(sib_b, by = "key") %>%
    left_join(ol_only_b, by = "key") %>%
    mutate(Other_Alloc_Barcodes = ifelse(is.na(OL_only), Other_Alloc_Barcodes,
                                         ifelse(is.na(Other_Alloc_Barcodes), OL_only,
                                                paste(Other_Alloc_Barcodes, OL_only, sep = ", ")))) %>%
    select(-OL_only) %>%
    mutate(
      OL_match  = coalesce(OL_match, "Not in Original List"),
      APOE      = sanger_cat(APOE_raw),
      ABCA7     = sanger_cat(ABCA7_raw),
      GSA_Stage = case_when(
        in_gsa & is_returned(GSA_Back) ~ "Returned",
        in_gsa                         ~ "At GSA",
        in_aw                          ~ "Awaiting GSA",
        TRUE                           ~ "Not sent"
      ),
      WGS_Stage = case_when(
        in_wgs & is_returned(WGS_Back) ~ "Returned",
        in_wgs & (toupper(coalesce(WGS_Back, "")) == "NEVER" |
                    grepl("USUHS", coalesce(WGS_Type, ""), fixed = TRUE)) ~ "Sent to USUHS",
        in_wgs                         ~ "At WGS",
        TRUE                           ~ "Not sent"
      )
    ) %>%
    arrange(Barcode)
}

# Gene names compared without spaces/punctuation/case ("ApoE ", "ABCA-7")
norm_gene <- function(x) {
  y <- toupper(gsub("[^A-Za-z0-9]", "", coalesce(as.character(x), "")))
  y[y == ""] <- NA_character_
  y
}

# Sanger_Rerun_List: identical rows (same barcode, SampleID, aliquot, allocation,
# gene and date) are counted once; gene spelling is normalised.
clean_rerun_log <- function(rlog) {
  if (is.null(rlog) || !nrow(rlog)) return(rlog)
  cols <- intersect(c("Barcode", "SampleID", "AliquotID", "Allocation", "Gene", "Date"), names(rlog))
  out <- rlog[!duplicated(rlog[cols]), , drop = FALSE]
  if ("Gene" %in% names(out)) out$Gene <- norm_gene(out$Gene)
  out
}

# Sanger_Rerun_List rows -> one row per rerun with the sample it counts for.
# A row with a barcode but no SampleID counts for the barcode's owner (final
# reports / logs / backup, where they agree); a row with neither is skipped.
# Blank Gene counts as APOE.
rerun_log_rows <- function(rlog, bt) {
  # who owns each barcode (final reports / logs / backup, where they agree)
  owner <- bt %>%
    filter(valid_bc(Barcode), nz(SampleID), is.na(ID_conflict)) %>%
    transmute(bk = norm_id(Barcode), owner = norm_sid(SampleID)) %>%
    distinct(bk, .keep_all = TRUE)
  
  # A rerun row with a barcode but no SampleID is counted for the barcode's owner;
  # a row with neither can't be counted for anyone.
  rl <- if (!is.null(rlog) && nrow(rlog) && "SampleID" %in% names(rlog)) {
    rlog %>%
      mutate(bk = norm_id(col_or_na(., "Barcode"))) %>%
      left_join(owner, by = "bk") %>%
      mutate(sid_key = ifelse(nz(SampleID), norm_sid(SampleID), owner)) %>%
      filter(nz(sid_key)) %>%
      transmute(sid_key,
                g = toupper(trimws(coalesce(col_or_na(., "Gene"), ""))),
                g = ifelse(g == "", "APOE", g),
                Barcode = col_or_na(., "Barcode"),
                Allocation = col_or_na(., "Allocation"),
                Date = col_or_na(., "Date"))
  } else tibble(sid_key = character(), g = character(), Barcode = character(),
                Allocation = character(), Date = character())
  list(owner = owner, rl = rl)
}

# Every allocation tied to a sample: its Original List rows, the allocation of
# each of its barcodes, and the allocations in Sanger_Rerun_List. Long: sid_key, a.
sample_allocs <- function(ol, bt, rl, req = NULL, extra = NULL) {
  bind_rows(
    if (!is.null(extra) && nrow(extra)) extra %>% transmute(sid_key, a = Allocation),
    if (!is.null(req) && nrow(req)) req %>% transmute(sid_key, a = Allocation),
    if (!is.null(ol) && nrow(ol)) ol %>% filter(nz(Sample_ID)) %>%
      transmute(sid_key = norm_sid(Sample_ID), a = col_or_na(., "AllocationID")),
    if (!is.null(bt) && nrow(bt)) bt %>% filter(nz(SampleID), valid_bc(Barcode)) %>%
      transmute(sid_key = norm_sid(SampleID), a = Allocation),
    if (!is.null(rl) && nrow(rl)) rl %>% transmute(sid_key, a = Allocation),
    tibble(sid_key = character(), a = character())
  ) %>%
    separate_rows(a, sep = ",\\s*") %>%
    mutate(a = alloc_base(a)) %>%
    filter(nz(a), !norm_id(a) %in% PLACEHOLDER_IDS, !NON_ALLOCATION(a)) %>%
    distinct()
}

SANGER_COLS <- c(APOE = "APOESanger(Result)", ABCA7 = "ABCA7Sanger(Result)")

# Runs per barcode, per gene = every row in Sanger_Rerun_List (one run each)
# + 1 for the sample's first run, which is never in the list: everything was
# run once before it reached the list, even if not by you. That +1 goes on the
# barcode of the sample's EARLIEST allocation (lowest allocation number, leaving
# out USUHS / NCRAD / CU-only tubes that were never received), preferring a
# tube of that allocation that's in the list; with no tube for it, on the
# earliest-dated row's barcode (or UNKNOWN).
#   20-2611260: 38966 (CU only) + 38897 -> FA21426758 = 1 row + first run = 2
#   20-2521049: 37858 + 38754 -> FS98147807 1 + 1 = 2, FA21418440 2 -> 4 runs
# Rows with no barcode are counted as "UNKNOWN" for the sample.
run_counts <- function(rl, ol = NULL, bt = NULL) {
  empty <- list(bc  = tibble(sid_key = character(), g = character(), bk = character(), runs = integer()),
                unk = tibble(sid_key = character(), g = character(), runs = integer()))
  if (is.null(rl)) rl <- tibble(sid_key = character(), g = character(), Barcode = character(),
                                Allocation = character(), Date = character())
  r <- rl %>% mutate(bk = if_else(valid_bc(Barcode), norm_id(Barcode), NA_character_), ord = row_number())
  bc  <- r %>% filter(!is.na(bk)) %>% count(sid_key, g, bk, name = "runs")
  unk <- r %>% filter(is.na(bk))  %>% count(sid_key, g, name = "runs")
  if (is.null(ol) || is.null(bt) || !nrow(ol)) {
    if (!nrow(r)) return(empty)
    return(list(bc = bc, unk = unk))
  }
  # every sample + gene that was run at all: a result in Original List, or a row
  ol2 <- ol %>% filter(nz(Sample_ID)) %>% mutate(sid_key = norm_sid(Sample_ID))
  ran <- bind_rows(
    bind_rows(lapply(names(SANGER_COLS), function(gn)
      ol2 %>% transmute(sid_key, g = gn, cat = sanger_cat(col_or_na(ol2, SANGER_COLS[[gn]]))))) %>%
      filter(!is.na(cat)) %>% select(sid_key, g),
    r %>% select(sid_key, g)) %>% distinct()
  if (!nrow(ran)) return(empty)
  # the sample's tubes and their allocations (tubes never received left out)
  bc_alloc <- bt %>% filter(valid_bc(Barcode)) %>%
    transmute(bk = key, ba = alloc_base(sub(",.*$", "", Allocation))) %>% distinct(bk, .keep_all = TRUE)
  rows_bc <- r %>% filter(!is.na(bk)) %>% left_join(bc_alloc, by = "bk") %>%
    mutate(a = alloc_base(coalesce(clean_alloc(Allocation), ba)))
  cand <- bind_rows(
    bt %>% filter(nz(SampleID), valid_bc(Barcode)) %>%
      transmute(sid_key = norm_sid(SampleID), bk = key, a = alloc_base(sub(",.*$", "", Allocation))),
    rows_bc %>% transmute(sid_key, bk, a)) %>%
    filter(nz(a)) %>% distinct() %>%
    anti_join(sent_map(ol), by = c("sid_key", "a")) %>%
    mutate(n = suppressWarnings(as.numeric(gsub("[^0-9]", "", a)))) %>%
    filter(!is.na(n))
  earliest <- cand %>% group_by(sid_key) %>% filter(n == min(n)) %>% ungroup()
  in_list <- r %>% filter(!is.na(bk)) %>% distinct(sid_key, g, bk) %>% mutate(listed = TRUE)
  pick <- ran %>%
    inner_join(earliest %>% select(sid_key, bk), by = "sid_key", relationship = "many-to-many") %>%
    left_join(in_list, by = c("sid_key", "g", "bk")) %>%
    arrange(sid_key, g, desc(coalesce(listed, FALSE)), bk) %>%
    group_by(sid_key, g) %>% slice(1) %>% ungroup() %>%
    transmute(sid_key, g, bk, runs = 1L)
  # no tube known for any allocation: the earliest-dated row's barcode, else UNKNOWN
  rest <- ran %>% anti_join(pick, by = c("sid_key", "g"))
  first_row <- r %>% semi_join(rest, by = c("sid_key", "g")) %>%
    arrange(sid_key, g, is.na(as_day(Date)), as_day(Date), ord) %>%
    group_by(sid_key, g) %>% slice(1) %>% ungroup()
  rest_bc  <- first_row %>% filter(!is.na(bk)) %>% transmute(sid_key, g, bk, runs = 1L)
  rest_unk <- rest %>% anti_join(rest_bc, by = c("sid_key", "g")) %>% transmute(sid_key, g, runs = 1L)
  list(bc  = bind_rows(bc, pick, rest_bc) %>% group_by(sid_key, g, bk) %>% summarise(runs = sum(runs), .groups = "drop"),
       unk = bind_rows(unk, rest_unk) %>% group_by(sid_key, g) %>% summarise(runs = sum(runs), .groups = "drop"))
}

# "APOE / ABCA7": how many times each barcode appears in Sanger_Rerun_List,
# keyed by normalised barcode
barcode_rerun_summary <- function(rl) {
  rl <- rl %>% filter(valid_bc(Barcode))
  if (!nrow(rl)) return(setNames(character(), character()))
  x <- rl %>% mutate(bk = norm_id(Barcode)) %>% count(bk, g) %>%
    tidyr::pivot_wider(names_from = g, values_from = n, values_fill = 0L)
  apoe  <- if ("APOE"  %in% names(x)) x$APOE  else 0L
  abca7 <- if ("ABCA7" %in% names(x)) x$ABCA7 else 0L
  setNames(paste(apoe, "/", abca7), x$bk)
}

# Sanger rerun list, one row per SampleID per gene, built from Original List.
# A sample is listed when it has a plain Fail for that gene and
#   - no Pass for that gene in any Original List row (any allocation),
#   - its most recent result for that gene isn't "Fail, Dropped" (dropped, not repeated),
#   - not every one of its rows has a Status starting with "Dropped".
# Samples in the Requested Reruns table are listed too, even after a Pass.
# Times Run = how many times the sample appears in Sanger_Rerun_List for that
# gene (at least 1). Rows with a blank Gene count as APOE.
last_by <- function(x, key) {
  if (!length(x)) return(NA_character_)
  x[order(key, na.last = FALSE)][length(x)]
}

build_reruns <- function(ol, bt, rlog = NULL, req = NULL, rx = NULL) {
  if (is.null(req)) req <- read_requested(NULL, character())
  # reruns logged per SampleID + gene
  parsed <- rerun_log_rows(rlog, bt)
  owner  <- parsed$owner
  rl     <- parsed$rl
  rc     <- run_counts(rl, ol, bt)
  # dates the sample appears in Sanger_Rerun_List, per gene
  rdates <- if (nrow(rl)) rl %>% group_by(sid_key, g) %>%
    summarise(`Rerun dates` = date_list(Date), .groups = "drop")
  else tibble(sid_key = character(), g = character(), `Rerun dates` = character())
  rtotal <- bind_rows(rc$bc %>% select(sid_key, g, runs), rc$unk) %>%
    group_by(sid_key, g) %>% summarise(total = sum(runs), .groups = "drop")
  
  ol2 <- ol %>%
    filter(nz(Sample_ID)) %>%
    mutate(sid_key  = norm_sid(Sample_ID),
           AllocationID = col_or_na(., "AllocationID"),
           Received = col_or_na(., "AllocationReceived"),
           Substudy = col_or_na(., "Substudy"))
  
  # Where each tube is now, one place per tube (first match wins):
  # out at GSA / WGS / USUHS, waiting for GSA, its backup plate + position,
  # or blank when none of the sheets says where it is. Each place is labelled
  # with its barcode; a sample's tubes are separated by "; ".
  tubes <- bt %>%
    filter(nz(SampleID)) %>%
    mutate(sid_key = norm_sid(SampleID),
           where = case_when(
             !is.na(Sent_To)                ~ NEVER_RECEIVED(Sent_To),
             GSA_Stage == "At GSA"          ~ "At GSA",
             WGS_Stage == "At WGS"          ~ "At WGS",
             WGS_Stage == "Sent to USUHS"   ~ "USUHS",
             GSA_Stage == "Awaiting GSA"    ~ "Awaiting GSA",
             Backup_Stage %in% "In backup"  ~ paste("Backup", Backup_Plate, Backup_Pos),
             TRUE                           ~ NA_character_),
           # say which tube is where: "FA21418440: Backup BackupInventory2025_p155 F10"
           where = ifelse(is.na(where), NA_character_,
                          ifelse(valid_bc(Barcode), paste0(Barcode, ": ", where), where))) %>%
    group_by(sid_key) %>%
    summarise(Location = collapse_u(where, "; "), .groups = "drop")
  
  # every barcode of the sample (its tubes + any in the rerun log), each with how
  # many times it appears in Sanger_Rerun_List for that gene: "FA12345678 (2)"
  tube_bc <- bt %>%
    filter(nz(SampleID), valid_bc(Barcode)) %>%
    transmute(sid_key = norm_sid(SampleID), Barcode, Allocation)
  # A barcode is one tube, so it's listed under the sample that owns it per the
  # final reports / logs / backup inventory, even if a Sanger_Rerun_List row
  # gives another SampleID. (n) = runs on that barcode for this gene (see run_counts).
  barcode_labels <- function(gene) {
    rg <- rl %>% filter(g == gene, valid_bc(Barcode)) %>% mutate(bk = norm_id(Barcode))
    n_bc <- rc$bc %>% filter(g == gene) %>% group_by(bk) %>% summarise(n = sum(runs), .groups = "drop")
    rg_attr <- rg %>%
      left_join(owner, by = "bk") %>%
      transmute(sid_key = coalesce(owner, sid_key), Barcode, bk)
    known <- bind_rows(tube_bc %>% transmute(sid_key, Barcode, bk = norm_id(Barcode)), rg_attr) %>%
      distinct(sid_key, bk, .keep_all = TRUE) %>%
      left_join(n_bc, by = "bk") %>%
      transmute(sid_key, lbl = paste0(Barcode, " (", coalesce(n, 0L), ")"))
    # runs logged with a blank barcode: "UNKNOWN (n)"
    unknown <- rc$unk %>%
      filter(g == gene) %>%
      transmute(sid_key, lbl = paste0("UNKNOWN (", runs, ")"))
    bind_rows(known, unknown) %>%
      group_by(sid_key) %>%
      summarise(Barcodes = paste(lbl, collapse = ", "), .groups = "drop")
  }
  
  # every allocation tied to the sample: its Original List rows, the allocation
  # of each of its barcodes, and the allocations in Sanger_Rerun_List
  sm <- sent_map(ol)
  all_allocs <- sample_allocs(ol, bt, rl, req, rx) %>%
    left_join(sm, by = c("sid_key", "a")) %>%
    group_by(sid_key) %>%
    summarise(Allocations = {
      o <- alloc_sort(a)
      d <- dest[match(o, a)]
      paste(ifelse(is.na(d), o, paste0(o, " (", d, " only)")), collapse = ", ")
    }, .groups = "drop")
  
  genes <- c(APOE = "APOESanger(Result)", ABCA7 = "ABCA7Sanger(Result)")
  out <- lapply(names(genes), function(g) {
    d <- ol2 %>% mutate(raw = col_or_na(ol2, genes[[g]]), cat = sanger_cat(raw)) %>% filter(!is.na(cat))
    rq <- req %>% filter(g == !!g) %>% distinct(sid_key, .keep_all = TRUE)
    # a request whose Result is Pass is done: it comes off the list
    rq_done <- rq$sid_key[req_closed(rq$Result)]
    rq <- rq %>% filter(!sid_key %in% rq_done)
    s <- if (nrow(d)) d %>%
      group_by(sid_key) %>%
      summarise(
        any_pass  = any(cat == "Pass"),
        # the most recent result decides: a "Fail, Dropped" followed by a plain
        # Fail in a later allocation (e.g. after reallocation) still needs a redo
        last_cat  = last_by(cat, Received),
        all_dropd = all(grepl("^Dropped", coalesce(Status, ""))),
        n_fail    = sum(cat == "Fail"),
        SampleID  = first(Sample_ID),
        `Failed in allocations` = collapse_u(AllocationID[cat == "Fail"]),
        `Sanger result`  = last_by(raw[cat == "Fail"], Received[cat == "Fail"]),
        last_raw  = last_by(raw, Received),
        Substudy  = collapse_u(Substudy),
        .groups = "drop"
      ) else tibble(sid_key = character(), any_pass = logical(), last_cat = character(),
                    all_dropd = logical(), n_fail = integer(), SampleID = character(),
                    `Failed in allocations` = character(), `Sanger result` = character(),
                    last_raw = character(), Substudy = character())
    # Requested Reruns that Original List doesn't put on the list by itself
    add <- rq %>% filter(!sid_key %in% s$sid_key) %>%
      transmute(sid_key, any_pass = FALSE, last_cat = NA_character_, all_dropd = FALSE,
                n_fail = 0L, SampleID)
    s <- bind_rows(s, add) %>%
      mutate(requested = sid_key %in% rq$sid_key) %>%
      # an open request keeps it on the list even after a Pass in Original List
      # (e.g. WGS discrepancy); otherwise a Pass anywhere takes it off
      filter(!sid_key %in% rq_done,
             requested | (n_fail > 0 & !any_pass & coalesce(last_cat, "") != "Fail, dropped" & !all_dropd)) %>%
      mutate(`Sanger result` = case_when(nz(`Sanger result`) ~ `Sanger result`,
                                         nz(last_raw)        ~ last_raw,
                                         TRUE                ~ "Requested"))
    if (!nrow(s)) return(NULL)
    s %>%
      mutate(Gene = g, Requested = if_else(requested, CHECK_YES, NA_character_)) %>%
      left_join(rq %>% select(sid_key, Requester, Note), by = "sid_key") %>%
      left_join(rtotal %>% filter(g == !!g) %>% select(sid_key, total), by = "sid_key") %>%
      mutate(`Times Run` = pmax(1L, coalesce(total, 0L))) %>%
      left_join(rdates %>% filter(g == !!g) %>% select(sid_key, `Rerun dates`), by = "sid_key") %>%
      left_join(barcode_labels(g), by = "sid_key")
  })
  bind_rows(out) %>%
    left_join(tubes, by = "sid_key") %>%
    left_join(all_allocs, by = "sid_key") %>%
    # "Failed in allocations" is kept for the NEED TO REDO counts, not displayed
    transmute(Gene, SampleID, `Times Run`, `Rerun dates`, Allocations, `Failed in allocations`,
              `Sanger result`, Substudy, Barcodes,
              Location, Requested, Requester, Note) %>%
    arrange(Gene, desc(`Times Run`), SampleID)
}

# Sanger_Rerun_List rows whose barcode doesn't match its SampleID: the barcode
# belongs to another sample per the final reports / backup, or the same barcode
# is logged under more than one SampleID in the list.
rerun_id_issues <- function(rlog, bt) {
  empty <- tibble(Barcode = character(), `Rerun list SampleID` = character(), AliquotID = character(),
                  `SampleID per AliquotID` = character(),
                  `Owner per final reports/backup` = character(), Problem = character())
  if (is.null(rlog) || !nrow(rlog) || !all(c("Barcode", "SampleID") %in% names(rlog))) return(empty)
  rl <- rlog %>%
    filter(valid_bc(Barcode), nz(SampleID)) %>%
    transmute(Barcode, SampleID, bk = norm_id(Barcode), sk = norm_sid(SampleID),
              AliquotID = col_or_na(., "AliquotID"),
              Gene = col_or_na(., "Gene"), Date = col_or_na(., "Date"))
  if (!nrow(rl)) return(empty)
  own <- bt %>% filter(valid_bc(Barcode), nz(SampleID)) %>%
    transmute(bk = norm_id(Barcode), owner_sid = SampleID, ok = norm_sid(SampleID)) %>%
    distinct(bk, .keep_all = TRUE)
  multi <- rl %>% group_by(bk) %>% summarise(n_sid = n_distinct(sk), .groups = "drop")
  rl %>%
    left_join(own, by = "bk") %>%
    left_join(multi, by = "bk") %>%
    mutate(Problem = case_when(
      !is.na(ok) & ok != sk ~ "Barcode belongs to another sample",
      n_sid > 1             ~ "Same barcode under more than one SampleID",
      TRUE                  ~ NA_character_)) %>%
    filter(!is.na(Problem)) %>%
    # the aliquot ID starts with its SampleID (202408352-35 -> 20-2408352), so it
    # shows which of Barcode / SampleID is the odd one out in that row
    mutate(aliq_sid = ifelse(grepl("^[0-9]{9}", norm_id(AliquotID)),
                             paste0(substr(norm_id(AliquotID), 1, 2), "-", substr(norm_id(AliquotID), 3, 9)),
                             NA_character_)) %>%
    transmute(Barcode, `Rerun list SampleID` = SampleID, AliquotID,
              `SampleID per AliquotID` = aliq_sid,
              `Owner per final reports/backup` = owner_sid, Gene, Date, Problem) %>%
    arrange(Barcode, Date)
}

# Original List vs final reports (ALL_FinalReports), per allocation, for
# allocations that are in both: samples in a final report with no Original
# List row, Original List samples (not Dropped) with no final report row,
# Dropped samples that do have a tube, and repeated Original List rows.
# Allocations in Original List with no final report at all are listed once.
ol_fr_check <- function(ol, afr) {
  empty <- tibble(Allocation = character(), SampleID = character(), Problem = character(),
                  `Original List status` = character(), Barcode = character())
  if (is.null(ol) || !nrow(ol) || is.null(afr) || !nrow(afr)) return(empty)
  o <- ol %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), SampleID = Sample_ID,
              a = alloc_base(col_or_na(., "AllocationID")), status = col_or_na(., "Status")) %>%
    filter(nz(a), !NON_ALLOCATION(a), !norm_id(a) %in% PLACEHOLDER_IDS) %>%
    mutate(dropped = grepl("^\\s*Dropped", coalesce(status, "")))
  f <- afr %>% filter(nz(SampleID)) %>%
    transmute(sid_key = norm_sid(SampleID), SampleID, a = alloc_base(order_id),
              Barcode = col_or_na(., "barcode_2d")) %>%
    filter(nz(a)) %>%
    group_by(sid_key, a) %>%
    summarise(SampleID = first(SampleID), n_bc = n_distinct(norm_id(Barcode[valid_bc(Barcode)])),
              Barcode = collapse_u(Barcode), .groups = "drop")
  both <- intersect(unique(o$a), unique(f$a))
  ok <- o %>% group_by(sid_key, a) %>%
    summarise(SampleID = first(SampleID), status = collapse_u(status, " | "),
              dropped = all(dropped), n = n(), .groups = "drop")
  miss <- f %>% filter(a %in% both) %>% anti_join(ok, by = c("sid_key", "a")) %>%
    transmute(Allocation = a, SampleID, Problem = "In the final report, missing from Original List",
              `Original List status` = NA_character_, Barcode)
  extra <- ok %>% filter(a %in% both, !dropped) %>% anti_join(f, by = c("sid_key", "a")) %>%
    transmute(Allocation = a, SampleID, Problem = "In Original List (not dropped), missing from the final report",
              `Original List status` = status, Barcode = NA_character_)
  drop_tube <- ok %>% filter(a %in% both, dropped) %>% inner_join(f %>% select(sid_key, a, Barcode), by = c("sid_key", "a")) %>%
    transmute(Allocation = a, SampleID, Problem = "Dropped in Original List, but has a tube in the final report",
              `Original List status` = status, Barcode)
  # the same sample more than once in an allocation is fine when each row has
  # its own barcode in the final report; flag only rows beyond the barcodes
  dups <- ok %>% filter(n > 1) %>%
    left_join(f %>% select(sid_key, a, n_bc, Barcode), by = c("sid_key", "a")) %>%
    filter(n > coalesce(n_bc, 0L)) %>%
    transmute(Allocation = a, SampleID,
              Problem = sprintf("Listed %d times in Original List but %d barcode(s) in the final report", n, coalesce(n_bc, 0L)),
              `Original List status` = status, Barcode)
  no_fr <- o %>% filter(!a %in% unique(f$a)) %>% group_by(a) %>%
    summarise(n = sum(!dropped), .groups = "drop") %>% filter(n > 0) %>%
    transmute(Allocation = a, SampleID = NA_character_,
              Problem = sprintf("Allocation has no final report yet (%d samples not dropped)", n),
              `Original List status` = NA_character_, Barcode = NA_character_)
  bind_rows(miss, extra, drop_tube, dups, no_fr) %>% arrange(Allocation, Problem, SampleID)
}

# Samples dropped in 2 or more allocations (Original List Status starting with
# "Dropped"; the "D-ropped" marker doesn't count), with each drop's allocation,
# reason and date, and the sample's other allocations.
dropped_repeat <- function(ol, recv_map = NULL) {
  empty <- tibble(SampleID = character(), `Times dropped` = integer(), `Dropped in` = character(),
                  `Other allocations` = character(), `Latest status` = character(),
                  APOE = character(), ABCA7 = character(), Substudy = character())
  if (is.null(ol) || !nrow(ol)) return(empty)
  o <- ol %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), SampleID = Sample_ID,
              a = alloc_base(col_or_na(., "AllocationID")), st = col_or_na(., "Status"),
              when = as_day(col_or_na(., "AllocationReceived")),
              apoe = col_or_na(., SANGER_COLS[["APOE"]]), abca7 = col_or_na(., SANGER_COLS[["ABCA7"]]),
              sub = col_or_na(., "Substudy")) %>%
    filter(nz(a)) %>%
    # no date on the row: the allocation's received date from elsewhere in the workbook
    mutate(when = if (length(recv_map)) coalesce(when, as_day(unname(recv_map[a]))) else when) %>%
    mutate(dropped = grepl("^\\s*Dropped", coalesce(st, "")))
  rep_s <- o %>% filter(dropped) %>% group_by(sid_key) %>% filter(n_distinct(a) >= 2) %>% ungroup()
  if (!nrow(rep_s)) return(empty)
  lab <- function(a, st, when) paste0(a, " (", trimws(sub("^\\s*Dropped\\s*:?\\s*", "", coalesce(st, ""))),
                                      ifelse(is.na(when), "", paste0(", ", format(when, "%m/%d/%Y"))), ")")
  drops <- rep_s %>% arrange(sid_key, when) %>% group_by(sid_key) %>%
    summarise(SampleID = first(SampleID), `Times dropped` = n_distinct(a),
              `Dropped in` = paste(unique(lab(a, st, when)), collapse = "; "), .groups = "drop")
  others <- o %>% filter(sid_key %in% drops$sid_key, !dropped) %>% arrange(sid_key, when) %>%
    group_by(sid_key) %>%
    summarise(`Other allocations` = paste(unique(paste0(a, ifelse(nz(st), paste0(" (", st, ")"), ""))), collapse = "; "),
              .groups = "drop")
  latest <- o %>% filter(sid_key %in% drops$sid_key) %>% group_by(sid_key) %>%
    summarise(`Latest status` = last_by(st, when), APOE = last_by(apoe[nz(apoe)], when[nz(apoe)]),
              ABCA7 = last_by(abca7[nz(abca7)], when[nz(abca7)]), Substudy = collapse_u(sub), .groups = "drop")
  drops %>% left_join(others, by = "sid_key") %>% left_join(latest, by = "sid_key") %>%
    arrange(desc(`Times dropped`), SampleID) %>%
    select(SampleID, `Times dropped`, `Dropped in`, `Other allocations`, `Latest status`, APOE, ABCA7, Substudy)
}

# Overview, built by the app ("Automatic" sub-tab): one row per allocation in
# Original List or ALL_FinalReports ("_#" folded into the parent), counted with
# the same rules as the rest of the app. Columns you fill in by hand on
# Overview_Tracker_2.0 (Comment, Sample #, Manifest, STAPLe, Genolist) are
# copied from that sheet by allocation number.
build_overview_auto <- function(ol, afr, gsa, wgs, ov, recv_map) {
  okA <- function(a) nz(a) & !NON_ALLOCATION(a) & !norm_id(a) %in% PLACEHOLDER_IDS
  o <- ol %>% filter(nz(Sample_ID)) %>%
    transmute(a = alloc_base(col_or_na(., "AllocationID")),
              apoe = sanger_cat(col_or_na(., SANGER_COLS[["APOE"]])),
              abca7 = sanger_cat(col_or_na(., SANGER_COLS[["ABCA7"]])),
              sex = col_or_na(., "Sex"), sub = col_or_na(., "Substudy"), eth = col_or_na(., "Ethnicity")) %>%
    filter(okA(a))
  f <- afr %>% transmute(a = alloc_base(order_id)) %>% filter(okA(a))
  day_counts <- function(x) {
    x <- x[nz(x)]
    if (!length(x)) return(NA_character_)
    d <- as_day(x); lab <- ifelse(is.na(d), x, format(d, "%m/%d/%Y"))
    t <- table(lab); k <- names(t)
    ord <- order(as_day(k), k, na.last = TRUE)
    paste0(k[ord], " (", as.integer(t[ord]), ")", collapse = ", ")
  }
  back_open <- function(x) !nz(x) | tolower(trimws(x)) %in% c("no", "n")
  g <- if (nrow(gsa)) gsa %>% transmute(a = alloc_base(col_or_na(., "Allocation")),
                                        h = col_or_na(., "Handoff DATE"), b = col_or_na(., "Handed Back")) %>%
    filter(okA(a)) else tibble(a = character(), h = character(), b = character())
  w <- if (nrow(wgs)) wgs %>% transmute(a = alloc_base(col_or_na(., "Allocation")),
                                        b = col_or_na(., "Handed Back")) %>%
    filter(okA(a)) else tibble(a = character(), b = character())
  allocs <- unique(c(o$a, f$a))
  if (!length(allocs)) return(tibble())
  base <- tibble(Allocation = allocs) %>%
    left_join(f %>% count(a, name = "Final #"), by = c("Allocation" = "a")) %>%
    left_join(o %>% group_by(a) %>% summarise(
      `Original #` = n(),
      `APOE Pass` = sum(apoe %in% "Pass"), `ABCA7 Pass` = sum(abca7 %in% "Pass"),
      # no call = plain Fail; "Fail, Dropped" is no longer part of the working allocation
      `APOE no call` = sum(apoe %in% "Fail"), `ABCA7 no call` = sum(abca7 %in% "Fail"),
      Sex = collapse_u(sex), Substudy = collapse_u(sub), Ethnicity = collapse_u(eth), .groups = "drop"),
      by = c("Allocation" = "a")) %>%
    left_join(g %>% group_by(a) %>% summarise(`Samples for GSA` = n(), `Samples at GSA` = sum(back_open(b)),
                                              `GSA handoff dates` = day_counts(h), `GSA returned` = day_counts(b),
                                              .groups = "drop"), by = c("Allocation" = "a")) %>%
    left_join(w %>% group_by(a) %>% summarise(`Samples for WGS` = n(), `Samples at WGS` = sum(back_open(b)),
                                              .groups = "drop"), by = c("Allocation" = "a")) %>%
    mutate(across(c(`Final #`, `Original #`, `APOE Pass`, `ABCA7 Pass`, `APOE no call`, `ABCA7 no call`,
                    `Samples for GSA`, `Samples at GSA`, `Samples for WGS`, `Samples at WGS`),
                  ~ coalesce(as.integer(.x), 0L)),
           Received = unname(recv_map[Allocation]))
  # hand-filled columns from Overview_Tracker_2.0
  pick <- function(pattern) {
    if (is.null(ov) || !nrow(ov)) return(NULL)
    j <- grep(pattern, names(ov), ignore.case = TRUE)
    if (length(j)) ov[[j[1]]] else NULL
  }
  if (!is.null(ov) && nrow(ov) && "Allocation" %in% names(ov)) {
    man <- tibble(Allocation = alloc_base(ov$Allocation))
    for (nm in c("Comment", "Sample #", "Manifest", "STAPLe", "Genolist")) {
      v <- pick(paste0("(^|\\W)", gsub("#", "\\#", nm, fixed = TRUE), "$"))
      man[[paste0(nm, " (manual)")]] <- if (is.null(v)) NA_character_ else as.character(v)
    }
    man <- man %>% filter(nz(Allocation)) %>% distinct(Allocation, .keep_all = TRUE)
    base <- base %>% left_join(man, by = "Allocation")
  }
  base %>%
    arrange(desc(as_day(Received)), desc(suppressWarnings(as.numeric(Allocation)))) %>%
    select(Allocation, `Final #`, `Original #`, Received, any_of("Comment (manual)"), any_of("Sample # (manual)"),
           `APOE Pass`, `ABCA7 Pass`, `APOE no call`, `ABCA7 no call`,
           `Samples for GSA`, `Samples at GSA`, Sex, Substudy,
           any_of(c("Manifest (manual)")), `GSA handoff dates`, any_of(c("STAPLe (manual)", "Genolist (manual)")),
           `GSA returned`, `Samples for WGS`, `Samples at WGS`, Ethnicity)
}

# ---- Aliquot -> SampleID -------------------------------------------------------
# Everything before the dash, as 2 digits + "-" + the rest: "202504587-47" ->
# "20-2504587", old 8-digit "19881212-04" -> "19-881212".
aliquot_to_sid <- function(a) {
  a <- trimws(a)
  m <- regmatches(a, regexec("^(\\d{2})(\\d{4,8})(-|$)", a))
  vapply(m, function(x) if (length(x) >= 3) paste0(x[2], "-", x[3]) else NA_character_, character(1))
}

# ---- PCR master mix (from PCR_CheatSheet) ---------------------------------------
PCR_RECIPE <- c(`Nuclease-Free Water` = 8.5, `Forward Primer (10 uM)` = 1.25, `Reverse Primer (10 uM)` = 1.25,
                `Q5 HiFi 2X MM` = 12.5, `DMSO 2%` = 0.5)
PCR_DNA_UL <- 1
CLEAN_RECIPE <- c(`PCR Product` = 5, `ExoSAP-IT` = 2)

# Volume <-> concentration units, as factors to ng and uL
MASS_U <- c(pg = 1e-3, ng = 1, ug = 1e3, mg = 1e6)
VOL_U  <- c(nL = 1e-3, uL = 1, mL = 1e3)
CONC_U <- c(`pg/uL` = 1e-3, `ng/uL` = 1, `ug/uL` = 1e3, `ng/mL` = 1e-3, `ug/mL` = 1, `mg/mL` = 1e3)   # in ng/uL
MOL_U  <- c(pmol = 1, nmol = 1e3, umol = 1e6)                       # in pmol
MCONC_U <- c(nM = 1e-3, uM = 1, mM = 1e3)                            # in pmol/uL
# mass <-> moles needs the molecular weight (g/mol): 1 pmol = MW / 1000 ng
to_ng   <- function(v, u, mw) if (u %in% names(MASS_U)) v * MASS_U[[u]] else v * MOL_U[[u]] * mw / 1000
from_ng <- function(ng, u, mw) if (u %in% names(MASS_U)) ng / MASS_U[[u]] else ng * 1000 / mw / MOL_U[[u]]
to_ngul   <- function(v, u, mw) if (u %in% names(CONC_U)) v * CONC_U[[u]] else v * MCONC_U[[u]] * mw / 1000
from_ngul <- function(x, u, mw) if (u %in% names(CONC_U)) x / CONC_U[[u]] else x * 1000 / mw / MCONC_U[[u]]
is_molar <- function(u) u %in% c(names(MOL_U), names(MCONC_U))

# ---- Submission prep (GSA / WGS plates) ---------------------------------------
# above each output table: download it as CSV. To copy and paste instead, use
# the table's own "Copy" button (copies every row, pastes straight into Excel).
sp_part_ui <- function(part) {
  if (!ALLOW_EXPORT) return(NULL)
  div(class = "d-flex align-items-center gap-2 mb-2",
      downloadButton(paste0("sp_csv_", part), "Download this as CSV", class = "btn-outline-primary btn-sm"),
      span(class = "status-line", "or use Copy on the table to paste into Excel (all rows)"))
}
# Same as GSA_WGS_SUBMISSION_PREP.xlsx: pasted barcodes, in that order, fill
# 96-well plates down the columns (A1, B1 ... H1, A2 ...), skipping the CEPH
# well(s); the plate letter goes up every plate. Tube data comes from
# ALL_FinalReports, Study / Substudy / Sex / Race / Ethnicity from Original List.
WELLS <- paste0(rep(LETTERS[1:8], 12), rep(1:12, each = 8))      # vertical order
prep_submission <- function(codes, d, gsa_name, wgs_name, letter = "A",
                            ceph_wells = "H12", ceph_id = "19-960561") {
  codes <- trimws(codes); codes <- codes[nzchar(codes)]
  if (!length(codes)) return(NULL)
  ceph <- toupper(trimws(unlist(strsplit(coalesce(ceph_wells, ""), "[,; ]+"))))
  ceph <- intersect(ceph, WELLS)
  free <- setdiff(WELLS, ceph)
  # tube rows from the final reports (first row per barcode)
  fr <- d$final_reports %>% filter(valid_bc(barcode_2d)) %>% mutate(k = norm_id(barcode_2d)) %>%
    distinct(k, .keep_all = TRUE)
  q <- tibble(n = seq_along(codes), Input = codes, k = norm_id(codes))
  rows <- q %>% left_join(fr, by = "k")
  sid <- coalesce(col_or_na(rows, "SampleID"),
                  vapply(col_or_na(rows, "child_aliquot"), function(a) {
                    m <- regmatches(a, regexec("^(\\d{2})(\\d{7})-", coalesce(a, "")))[[1]]
                    if (length(m) == 3) paste0(m[2], "-", m[3]) else NA_character_ }, character(1)))
  rows$SampleID <- sid
  rows$sid_key <- norm_sid(coalesce(sid, ""))
  # Original List details (the row of this tube's allocation first, else any row)
  ol <- d$orig %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
              Study = col_or_na(., "Study"), Substudy = col_or_na(., "Substudy"), Sex = col_or_na(., "Sex"),
              Race = col_or_na(., "Race"), Ethnicity = col_or_na(., "Ethnicity"), olStatus = col_or_na(., "Status"))
  rows$a <- alloc_base(col_or_na(rows, "order_id"))
  m1 <- rows %>% select(n, sid_key, a) %>% left_join(ol, by = c("sid_key", "a")) %>% distinct(n, .keep_all = TRUE)
  m2 <- ol %>% distinct(sid_key, .keep_all = TRUE) %>% select(-a)
  fill <- function(col) coalesce(m1[[col]], m2[[col]][match(rows$sid_key, m2$sid_key)])
  for (cc in c("Study", "Substudy", "Sex", "Race", "Ethnicity")) rows[[cc]] <- fill(cc)
  rows$olStatus <- m1$olStatus
  # warnings, from the app's own view of each tube
  bt <- d$bt %>% mutate(k = norm_id(coalesce(Barcode, ""))) %>% distinct(k, .keep_all = TRUE)
  b <- bt[match(rows$k, bt$k), , drop = FALSE]
  w <- character(nrow(rows))
  add <- function(cond, msg) w <<- ifelse(coalesce(cond, FALSE), ifelse(nzchar(w), paste(w, msg, sep = "; "), msg), w)
  add(is.na(rows$barcode_2d), "not in the final reports")
  add(duplicated(rows$k), "listed twice")
  add(b$GSA_Stage %in% c("At GSA", "Returned"), "already sent to GSA")
  add(b$WGS_Stage %in% c("At WGS", "Returned", "Sent to USUHS"), paste0("already sent to WGS (", coalesce(b$WGS_Stage, ""), ")"))
  add(!is.na(b$Sent_To), paste0(coalesce(b$Sent_To, ""), " only (never received)"))
  add(grepl("^\\s*Dropped", coalesce(rows$olStatus, "")), paste0("Original List: ", coalesce(rows$olStatus, "")))
  rows$Warning <- ifelse(nzchar(w), w, NA_character_)
  # wells
  per <- length(free)
  idx <- seq_len(nrow(rows)) - 1L
  plate_i <- idx %/% per
  rows$Letter <- vapply(plate_i, function(i) intToUtf8(utf8ToInt(toupper(substr(letter, 1, 1))) + i), character(1))
  rows$Well <- free[idx %% per + 1L]
  n_pl <- max(plate_i) + 1L
  cephs <- if (length(ceph)) tibble(
    Letter = rep(vapply(seq_len(n_pl) - 1L, function(i) intToUtf8(utf8ToInt(toupper(substr(letter, 1, 1))) + i), character(1)),
                 each = length(ceph)),
    Well = rep(ceph, n_pl)) else tibble(Letter = character(), Well = character())
  ord <- function(df) df %>% mutate(.o = match(Well, WELLS)) %>% arrange(Letter, .o) %>% select(-.o)
  num <- function(x) suppressWarnings(as.numeric(x))
  conc_ng <- num(col_or_na(rows, "concentration")) * 1000
  gsa <- bind_rows(
    tibble(`Aliquot ID` = col_or_na(rows, "child_aliquot"), `Sample ID` = rows$SampleID,
           Letter = rows$Letter, Well = rows$Well,
           `Amount (ug)` = num(col_or_na(rows, "amount")), `Conc (ng/uL)` = conc_ng,
           `Vol (uL)` = num(col_or_na(rows, "volume")), `Quant Method` = col_or_na(rows, "quantitation_method"),
           `Barcode 2D` = coalesce(col_or_na(rows, "barcode_2d"), rows$Input),
           Study = rows$Study, Substudy = rows$Substudy, Gender = rows$Sex, Comments = NA_character_),
    cephs %>% mutate(`Aliquot ID` = "CEPH", `Sample ID` = "CEPH")) %>%
    ord() %>%
    mutate(Plate = paste0(gsa_name, "_", Letter), POS = Well, Plate_POS = paste0(Plate, "_", POS)) %>%
    select(`Aliquot ID`, `Sample ID`, Plate, POS, Plate_POS, `Amount (ug)`, `Conc (ng/uL)`, `Vol (uL)`,
           `Quant Method`, `Barcode 2D`, Study, Substudy, Gender, Comments)
  wgs <- bind_rows(
    tibble(`External ID` = rows$SampleID, `HIHG ID` = col_or_na(rows, "child_aliquot"),
           `Sample Conc. (ng/ul)` = conc_ng, `Volume (uL)` = num(col_or_na(rows, "volume")),
           `Total Mass (ug)` = num(col_or_na(rows, "amount")), `Container Type` = "0.5mL FluidX",
           Letter = rows$Letter, Well = rows$Well, `2D BARCODE` = coalesce(col_or_na(rows, "barcode_2d"), rows$Input),
           Sex = rows$Sex, Race = rows$Race, Ethnicity = rows$Ethnicity, Comments = rows$Substudy),
    cephs %>% mutate(`External ID` = "-", `Container Type` = "0.5mL FluidX")) %>%   # CEPH well left empty
    ord() %>%
    mutate(`Container Name` = paste0(wgs_name, "_", Letter), `Sample Well Location_Vertical` = Well) %>%
    select(`External ID`, `HIHG ID`, `Sample Conc. (ng/ul)`, `Volume (uL)`, `Total Mass (ug)`, `Container Type`,
           `Container Name`, `Sample Well Location_Vertical`, `2D BARCODE`, Sex, Race, Ethnicity, Comments)
  staple_cols <- c("requestor", "parent_aliquot", "child_aliquot", "amount", "amount_unit", "concentration",
                   "concentration_unit", "volume", "volume_unit", "type", "allocation_type", "plate_id", "position",
                   "barcode_2d", "grow_code", "notes", "derivation", "local_ext_id", "sample_local_ext_id",
                   "quantitation_method", "avg_260_280", "avg_260_230")
  staple <- as_tibble(setNames(lapply(staple_cols, function(cc) col_or_na(rows, cc)), staple_cols)) %>%
    filter(!is.na(rows$barcode_2d))
  geno <- bind_rows(
    tibble(`Sample ID` = rows$SampleID, `Aliquot ID` = col_or_na(rows, "child_aliquot"),
           Letter = rows$Letter, Well = rows$Well, Notes = NA_character_),
    cephs %>% mutate(`Sample ID` = ceph_id, `Aliquot ID` = NA_character_, Notes = "CEPH")) %>%
    ord() %>% transmute(`Sample ID`, `Aliquot ID`, Plate = Letter, Pos = Well, Notes)
  check <- rows %>% transmute(`#` = n, Barcode = Input, SampleID, `Aliquot ID` = col_or_na(rows, "child_aliquot"),
                              Allocation = a, Plate = Letter, Well, Warning)
  list(check = check, gsa = gsa, wgs = wgs, staple = staple, geno = geno,
       n = nrow(rows), plates = n_pl, warn = sum(!is.na(rows$Warning)))
}

# Re-extraction, per gene, for samples on the Sanger Hiatus list, counted per
# barcode (each barcode is one tube). A barcode's fails = its runs for that gene
# (its rows in Sanger_Rerun_List + the sample's first run on its earliest
# allocation's tube; see run_counts), the same numbers as Times Run / Barcodes.
#  * "Meets criteria" = Recommended "Re-extract": the newest allocation's tube
#    was run 2+ times and the sample has 2+ allocations (see below).
# Rows with no barcode (UNKNOWN) can't be tied to a tube and aren't counted.
# A sample whose Status mentions "Re-extracted" is shown as Re-extracted.
build_reextract <- function(ol, bt, rlog, reruns, rx = NULL) {
  empty <- tibble(Gene = character(), sid_key = character(), SampleID = character(),
                  State = character(), `Fails per barcode` = character(), Recommended = character())
  if (is.null(reruns) || !nrow(reruns)) return(empty)
  rl <- rerun_log_rows(rlog, bt)$rl
  bc_alloc <- bt %>% filter(valid_bc(Barcode)) %>%
    transmute(bk = key, ba = alloc_base(sub(",.*$", "", Allocation))) %>% distinct(bk, .keep_all = TRUE)
  ol2 <- ol %>% filter(nz(Sample_ID)) %>%
    mutate(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
           st = paste(coalesce(col_or_na(., "Status"), ""), coalesce(col_or_na(., "Comments"), "")))
  # "Re-extracted" in Status or Comments (any row of the sample)
  done <- ol2 %>% group_by(sid_key) %>%
    # "Re-extracted" (past tense) only, so a note like "needs re-extraction" doesn't count
    summarise(done = any(grepl("re-?\\s*extracted", coalesce(st, ""), ignore.case = TRUE)), .groups = "drop")
  # samples in the Re-extracted Samples table are done too
  if (!is.null(rx) && nrow(rx))
    done <- bind_rows(done, tibble(sid_key = rx$sid_key, done = TRUE)) %>%
    group_by(sid_key) %>% summarise(done = any(done), .groups = "drop")
  in_ol <- ol2 %>% filter(nz(a)) %>% distinct(sid_key, a) %>% mutate(in_ol = TRUE)
  ol_fail <- bind_rows(lapply(names(SANGER_COLS), function(gn)
    ol2 %>% transmute(sid_key, a, g = gn, cat = sanger_cat(col_or_na(ol2, SANGER_COLS[[gn]]))))) %>%
    filter(nz(a)) %>%
    group_by(sid_key, a, g) %>%
    summarise(ol_fail = any(cat %in% c("Fail", "Fail, dropped")), .groups = "drop")
  
  # fails per barcode = its runs (first run + rows), same counter as Times Run;
  # the sample is on the Hiatus list for that gene, so every run failed
  rc <- run_counts(rl, ol, bt)
  bc_info <- bind_rows(
    bt %>% filter(valid_bc(Barcode)) %>% transmute(bk = key, Barcode, a = alloc_base(sub(",.*$", "", Allocation))),
    if (nrow(rl)) rl %>% filter(valid_bc(Barcode)) %>% mutate(bk = norm_id(Barcode)) %>%
      transmute(bk, Barcode, a = alloc_base(clean_alloc(Allocation)))) %>%
    filter(nz(bk)) %>% group_by(bk) %>%
    summarise(Barcode = first(Barcode), a = first(a[nz(a)]), .groups = "drop")
  on_list <- reruns %>% transmute(Gene, sid_key = norm_sid(SampleID), SampleID)
  per_bc <- rc$bc %>%
    semi_join(on_list %>% transmute(sid_key, g = Gene), by = c("sid_key", "g")) %>%
    left_join(bc_info, by = "bk") %>%
    mutate(Barcode = coalesce(Barcode, bk)) %>%
    left_join(in_ol, by = c("sid_key", "a")) %>%
    mutate(in_ol = coalesce(in_ol, FALSE), fails = runs) %>%
    filter(fails > 0)
  summ <- per_bc %>%
    group_by(sid_key, g) %>%
    summarise(
      n_bc2   = sum(fails >= 2),
      n_bc    = n_distinct(bk),
      total   = sum(fails),
      missing = any(!in_ol),
      `Fails per barcode` = paste0(Barcode, " [", coalesce(a, "?"),
                                   ifelse(in_ol, "", ", not in Original List"), "] (", fails, ")",
                                   collapse = ", "),
      .groups = "drop")
  # Recommended next step, from the sample's NEWEST allocation (highest
  # allocation number among tubes you received; USUHS / NCRAD / CU-only left out):
  #   Re-extract - newest allocation run 2+ times AND the sample has 2+ allocations
  #                (it was already reallocated, or both allocations ran twice)
  #   Reallocate - newest allocation run 2+ times, only one allocation
  #   Reallocate - also when the newest allocation's tube is at USUHS (can't be rerun)
  #   Rerun      - newest allocation run once (or not yet: a new tube)
  # Already re-extracted: "Run (re-extracted)" while the new extraction hasn't
  # been run, "Rerun (re-extracted)" after one run, then after two
  # "Re-extract again? (already re-extracted)".
  received <- bt %>% filter(nz(SampleID), valid_bc(Barcode)) %>%
    transmute(sid_key = norm_sid(SampleID), bk = key, a = alloc_base(sub(",.*$", "", Allocation)),
              usuhs = WGS_Stage %in% "Sent to USUHS", at_wgs = WGS_Stage %in% "At WGS") %>%
    filter(nz(a)) %>% anti_join(sent_map(ol), by = c("sid_key", "a"))
  # allocations whose tube(s) all went to USUHS (sent with WGS, never handed back)
  at_usuhs <- received %>% group_by(sid_key, a) %>%
    summarise(usuhs = all(usuhs), at_wgs = any(at_wgs), .groups = "drop")
  allocs_g <- bind_rows(
    on_list %>% transmute(sid_key, g = Gene) %>% distinct() %>%
      inner_join(received %>% select(sid_key, a), by = "sid_key", relationship = "many-to-many"),
    per_bc %>% filter(nz(a)) %>% transmute(sid_key, g, a)) %>%
    distinct() %>%
    mutate(n = suppressWarnings(as.numeric(gsub("[^0-9]", "", a)))) %>% filter(!is.na(n))
  runs_a <- per_bc %>% filter(nz(a)) %>% group_by(sid_key, g, a) %>% summarise(runs = sum(fails), .groups = "drop")
  rec <- allocs_g %>%
    left_join(runs_a, by = c("sid_key", "g", "a")) %>%
    mutate(runs = coalesce(runs, 0L)) %>%
    group_by(sid_key, g) %>%
    summarise(n_alloc = n_distinct(a), newest_runs = runs[which.max(n)], newest = a[which.max(n)],
              .groups = "drop") %>%
    left_join(at_usuhs %>% rename(newest = a, newest_usuhs = usuhs, newest_wgs = at_wgs), by = c("sid_key", "newest"))
  # "reallocation attempted": the sample's most recent Original List row (by
  # AllocationReceived) is Dropped, after an earlier Fail for that gene
  olr <- ol %>% filter(nz(Sample_ID)) %>%
    transmute(sid_key = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")),
              status = col_or_na(., "Status"), when = as_day(col_or_na(., "AllocationReceived")),
              APOE = sanger_cat(col_or_na(., SANGER_COLS[["APOE"]])),
              ABCA7 = sanger_cat(col_or_na(., SANGER_COLS[["ABCA7"]])))
  latest <- olr %>% filter(!is.na(when)) %>% group_by(sid_key) %>%
    filter(when == max(when)) %>% slice(1) %>% ungroup() %>%
    filter(grepl("^\\s*Dropped", coalesce(status, "")))
  attempted <- bind_rows(lapply(names(SANGER_COLS), function(gn) {
    earlier_fail <- olr %>% filter(.data[[gn]] %in% c("Fail", "Fail, dropped")) %>%
      select(sid_key, fail_when = when)
    latest %>% inner_join(earlier_fail, by = "sid_key", relationship = "many-to-many") %>%
      filter(is.na(fail_when) | fail_when < when) %>%
      distinct(sid_key, .keep_all = TRUE) %>%
      transmute(sid_key, Gene = gn,
                attempt = paste0("reallocation attempted: ", a, ", ",
                                 trimws(sub("^\\s*Dropped\\s*:?\\s*", "", status)), ", ",
                                 format(when, "%m/%d/%Y")))
  }))
  if (!nrow(attempted)) attempted <- tibble(sid_key = character(), Gene = character(), attempt = character())
  attempted$attempt <- gsub(", ,", ",", attempted$attempt, fixed = TRUE)
  
  # after a re-extraction, follow the NEW extraction: its allocation from the
  # Re-extracted Samples table, else the newest allocation
  rx_alloc <- if (!is.null(rx) && nrow(rx)) rx %>% transmute(sid_key, rx_a = alloc_base(Allocation))
  else tibble(sid_key = character(), rx_a = character())
  on_list %>%
    left_join(done, by = "sid_key") %>%
    left_join(summ, by = c("sid_key", "Gene" = "g")) %>%
    left_join(rec, by = c("sid_key", "Gene" = "g")) %>%
    left_join(rx_alloc, by = "sid_key") %>%
    left_join(attempted, by = c("sid_key", "Gene")) %>%
    left_join(runs_a %>% rename(Gene = g, rx_a = a, rx_runs = runs), by = c("sid_key", "Gene", "rx_a")) %>%
    mutate(new_runs = case_when(nz(rx_a) ~ coalesce(rx_runs, 0L),
                                TRUE     ~ coalesce(newest_runs, 0L))) %>%
    mutate(newest_runs = coalesce(newest_runs, coalesce(total, 1L)),
           n_alloc = coalesce(n_alloc, 1L),
           Recommended = case_when(coalesce(done, FALSE) & new_runs >= 2 ~ "Re-extract again? (already re-extracted)",
                                   coalesce(done, FALSE) & new_runs == 1 ~ "Rerun (re-extracted)",
                                   coalesce(done, FALSE)                ~ "Run (re-extracted)",
                                   newest_runs >= 2 & n_alloc >= 2      ~ "Re-extract",
                                   newest_runs >= 2                     ~ "Reallocate",
                                   # newest tube is at USUHS: can't be rerun
                                   coalesce(newest_usuhs, FALSE)        ~ "Reallocate",
                                   TRUE                                 ~ "Rerun"),
           Recommended = if_else(!is.na(attempt), paste0(Recommended, " (", attempt, ")"), Recommended),
           # newest tube is out at WGS (in house, not back yet)
           Recommended = if_else(coalesce(newest_wgs, FALSE),
                                 paste(Recommended, "(currently in WGS pipeline)"), Recommended),
           State = case_when(coalesce(done, FALSE)       ~ "Re-extracted",
                             startsWith(Recommended, "Re-extract ") | Recommended == "Re-extract" ~ "Meets criteria",
                             TRUE                        ~ NA_character_)) %>%
    select(Gene, sid_key, SampleID, State, `Fails per barcode`, Recommended)
}

# Every way a pasted ID can reach a barcode row.
build_lookup_index <- function(bt) {
  base <- bt %>% select(key)
  sids <- bt %>%
    select(key, afr_sid, old_sid, gsa_sid, wgs_sid, aw_sid, bk_sid) %>%
    pivot_longer(-key, values_to = "v") %>%
    filter(nz(v)) %>%
    separate_rows(v, sep = ",\\s*") %>%
    filter(!norm_id(v) %in% PLACEHOLDER_IDS) %>%
    transmute(q = norm_sid(v), key, via = "SampleID")
  aliq <- bt %>%
    select(key, afr_aliquot, old_aliquot, gsa_aliquot, wgs_aliquot, aw_aliquot, bk_aliquot) %>%
    pivot_longer(-key, values_to = "v") %>%
    filter(nz(v)) %>%
    separate_rows(v, sep = ",\\s*") %>%
    filter(!norm_id(v) %in% PLACEHOLDER_IDS) %>%
    transmute(q = norm_id(v), key, via = "Aliquot")
  bind_rows(base %>% transmute(q = key, key, via = "Barcode"), sids, aliq) %>%
    distinct(q, key, .keep_all = TRUE)
}

# Read from a private copy, not the live file. Excel (AutoSave/co-authoring)
# and OneDrive sync can hold the original open; copying usually still works,
# and it also forces OneDrive to download an online-only file. The first bytes
# tell us what we actually got.
snapshot_workbook <- function(path) {
  tmp <- tempfile("lab_tracker_", fileext = ".xlsx")
  ok <- suppressWarnings(file.copy(path, tmp, overwrite = TRUE))
  if (!isTRUE(ok) || !file.exists(tmp) || file.size(tmp) == 0) {
    stop("R couldn't copy the file. OneDrive may still be downloading it, or another program ",
         "has it locked. In File Explorer, right-click it \u2192 'Always keep on this device', ",
         "wait for the green check, then press Load again.")
  }
  sig <- readBin(tmp, "raw", 4)
  if (identical(sig, as.raw(c(0x50, 0x4B, 0x03, 0x04)))) return(tmp)  # normal .xlsx (zip)
  unlink(tmp)
  if (identical(sig, as.raw(c(0xD0, 0xCF, 0x11, 0xE0)))) {
    stop("This workbook is encrypted, either by a password or by a Microsoft sensitivity label ",
         "(e.g. 'Confidential' / 'Internal' from the university). No R package can read encrypted ",
         "Office files. In Excel check Home \u2192 Sensitivity and File \u2192 Info \u2192 Protect Workbook.")
  }
  stop(sprintf("The file isn't a readable .xlsx (it starts with bytes %s). If OneDrive shows a cloud icon ",
               "on it, make it available offline and try again.", paste(format(sig), collapse = " ")))
}

# paths: list(tracker = , backup = , old = ). The tracker is required; a problem
# with an optional workbook is reported but doesn't stop the load.
load_tracker <- function(paths) {
  path <- paths$tracker
  if (!nzchar(path) || !file.exists(path)) stop("File not found: ", normalizePath(path, mustWork = FALSE))
  mtime <- file.mtime(path)
  live  <- path
  path  <- snapshot_workbook(live)   # every read below uses the copy
  on.exit(unlink(path), add = TRUE)  # removed as soon as this load finishes
  sheets <- excel_sheets(path)
  afr <- read_simple(path, "ALL_FinalReports", "order_id")
  ol  <- drop_gsa_result(read_simple(path, "Original List", "File_Name"))
  gsa <- read_simple(path, "GSA", "Handoff DATE")
  wgs <- read_simple(path, "WGS", "Handoff DATE")
  aw  <- try_read(read_simple(path, "Awaiting_GSA", "Barcode"))
  
  extra_err <- list()
  read_extra <- function(p, reader, label) {
    if (!nzchar(p)) return(NULL)
    if (!file.exists(p)) {
      extra_err[[label]] <<- paste("File not found:", p)
      return(NULL)
    }
    snap <- tryCatch(snapshot_workbook(p), error = function(e) e)
    if (inherits(snap, "error")) {
      extra_err[[label]] <<- conditionMessage(snap)
      return(NULL)
    }
    on.exit(unlink(snap))
    tryCatch(reader(snap), error = function(e) {
      extra_err[[label]] <<- conditionMessage(e)
      NULL
    })
  }
  bk  <- read_extra(paths$backup, read_backup, "Backup inventory")
  old <- read_extra(paths$old, read_old_fr, "Old final reports")
  
  bt  <- build_barcode_table(afr, ol, gsa, wgs, aw, bk, old)
  ov  <- try_read(read_overview(path))
  rlog <- try_read(read_simple(path, "Sanger_Rerun_List", "Barcode"))
  rlog_raw_rows <- nrow(rlog)
  rlog <- clean_rerun_log(rlog)
  req <- tryCatch(read_requested(path, sheets), error = function(e)
    structure(read_requested(NULL, character()), load_error = conditionMessage(e)))
  rx_tbl <- tryCatch(read_reextracted(path, sheets), error = function(e) read_reextracted(NULL, character()))
  reruns_tbl <- tryCatch(build_reruns(ol, bt, rlog, req, rx_tbl), error = function(e)
    structure(data.frame(), load_error = conditionMessage(e)))
  # runs per barcode and per sample's UNKNOWN rows, "APOE / ABCA7"
  rc_all <- tryCatch(run_counts(rerun_log_rows(rlog, bt)$rl, ol, bt), error = function(e) run_counts(NULL))
  ab <- function(df, key) {
    if (!nrow(df)) return(setNames(character(), character()))
    x <- df %>% group_by(k = .data[[key]]) %>%
      summarise(a = sum(runs[g == "APOE"]), b = sum(runs[g == "ABCA7"]), .groups = "drop")
    setNames(paste(x$a, "/", x$b), x$k)
  }
  bc_runs <- ab(rc_all$bc, "bk")
  unk_map <- ab(rc_all$unk, "sid_key")
  # rerun dates per barcode (and per sample for rows with no barcode), per gene
  rdl <- tryCatch(rerun_log_rows(rlog, bt)$rl, error = function(e) NULL)
  date_map <- function(df, key) {
    if (is.null(df) || !nrow(df)) return(list(APOE = character(), ABCA7 = character()))
    x <- df %>% group_by(k = .data[[key]], g) %>% summarise(d = date_list(Date), .groups = "drop")
    list(APOE  = setNames(x$d[x$g == "APOE"],  x$k[x$g == "APOE"]),
         ABCA7 = setNames(x$d[x$g == "ABCA7"], x$k[x$g == "ABCA7"]))
  }
  bc_dates  <- date_map(if (!is.null(rdl)) rdl %>% filter(valid_bc(Barcode)) %>% mutate(bk = norm_id(Barcode)), "bk")
  unk_dates <- date_map(if (!is.null(rdl)) rdl %>% filter(!valid_bc(Barcode)), "sid_key")
  unknown_runs <- tibble(sid_key = names(unk_map), runs = unname(unk_map)) %>%
    mutate(SampleID = paste0(substr(sid_key, 1, 2), "-", substr(sid_key, 3, nchar(sid_key))))
  # per sample: reruns logged, and total runs = first run + reruns. The first run
  # counts for a gene if the sample has a result for it in Original List or has
  # been rerun for it (a rerun means there was a first run).
  run_maps <- tryCatch({
    r <- rerun_log_rows(rlog, bt)$rl
    rr_n <- r %>% group_by(sid_key) %>%
      summarise(ra = sum(g == "APOE"), rb = sum(g == "ABCA7"), .groups = "drop")
    first <- ol %>% filter(nz(Sample_ID)) %>%
      mutate(sid_key = norm_sid(Sample_ID),
             a_res = nz(col_or_na(., "APOESanger(Result)")),
             b_res = nz(col_or_na(., "ABCA7Sanger(Result)"))) %>%
      group_by(sid_key) %>%
      summarise(fa = any(a_res), fb = any(b_res), .groups = "drop")
    tot <- bind_rows(rc_all$bc %>% select(sid_key, g, runs), rc_all$unk) %>%
      group_by(sid_key) %>%
      summarise(ca = sum(runs[g == "APOE"]), cb = sum(runs[g == "ABCA7"]), .groups = "drop")
    x <- full_join(rr_n, first, by = "sid_key") %>%
      full_join(tot, by = "sid_key") %>%
      mutate(ra = coalesce(ra, 0L), rb = coalesce(rb, 0L),
             fa = coalesce(fa, FALSE) | ra > 0, fb = coalesce(fb, FALSE) | rb > 0,
             # all runs (first run + rows), at least 1 when there's a result
             ta = pmax(coalesce(ca, 0L), as.integer(fa)), tb = pmax(coalesce(cb, 0L), as.integer(fb)))
    list(reruns = setNames(paste(x$ra, "/", x$rb), x$sid_key),
         total  = setNames(paste(x$ta, "/", x$tb), x$sid_key))
  }, error = function(e) list(reruns = setNames(character(), character()),
                              total  = setNames(character(), character())))
  sample_runs <- run_maps$reruns
  # Requested Reruns rows to fix: unreadable Gene, odd SampleID, or a SampleID
  # that isn't in Original List, any barcode sheet or Sanger_Rerun_List
  req_issues <- tryCatch({
    known <- unique(c(norm_sid(ol$Sample_ID), norm_sid(bt$SampleID),
                      if (nrow(rlog) && "SampleID" %in% names(rlog)) norm_sid(rlog$SampleID)))
    req %>% distinct(Input, .keep_all = TRUE) %>%
      mutate(Problem = case_when(
        !is.na(Problem) ~ Problem,
        !sid_key %in% known ~ "Not in Original List, final reports or the rerun log (old sample, or a typo)",
        TRUE ~ NA_character_)) %>%
      filter(!is.na(Problem)) %>%
      transmute(SampleID = Input, Gene = GeneText, Requester, Problem)
  }, error = function(e) tibble(SampleID = character(), Gene = character(),
                                Requester = character(), Problem = character()))
  if (!is.null(attr(req, "load_error")))
    req_issues <- structure(req_issues, load_error = attr(req, "load_error"))
  # Reallocated = the sample is in two or more different allocations (Original
  # List rows + the allocations of its barcodes). A re-tube within the same
  # allocation doesn't count. Samples whose Status says "Reallocated" /
  # "Reallocation" are included too.
  alloc_pairs <- bind_rows(
    tibble(sid = norm_sid(ol$Sample_ID), a = col_or_na(ol, "AllocationID")),
    tibble(sid = norm_sid(bt$SampleID), a = bt$Allocation)
  ) %>%
    filter(nz(sid), !sid %in% PLACEHOLDER_IDS) %>%
    separate_rows(a, sep = ",\\s*") %>%
    mutate(a = alloc_base(a)) %>%
    filter(nz(a), !norm_id(a) %in% PLACEHOLDER_IDS, !NON_ALLOCATION(a)) %>%
    distinct()
  multi_alloc <- alloc_pairs %>% count(sid) %>% filter(n >= 2) %>% pull(sid)
  by_status <- norm_sid(ol$Sample_ID[nz(ol$Sample_ID) &
                                       grepl("reallo", coalesce(ol$Status, ""), ignore.case = TRUE)])
  realloc_sids <- unique(c(multi_alloc, by_status))
  bt$Reallocated <- ifelse(nz(bt$SampleID) & norm_sid(coalesce(bt$SampleID, "")) %in% realloc_sids,
                           CHECK_YES, NA_character_)
  rr <- unname(bc_runs[bt$key])
  bt$Reruns <- ifelse(valid_bc(bt$Barcode), coalesce(rr, "0 / 0"), NA_character_)
  
  # Date each allocation was received, from Original List "AllocationReceived".
  # It's the same for the whole allocation, so any row of that allocation can
  # supply it (most common date wins if they ever differ).
  recv <- tibble(a = alloc_base(col_or_na(ol, "AllocationID")), r = col_or_na(ol, "AllocationReceived")) %>%
    filter(nz(a), nz(r)) %>%
    group_by(a) %>%
    summarise(r = names(sort(table(r), decreasing = TRUE))[1], .groups = "drop")
  recv_map <- setNames(recv$r, recv$a)
  bt$Allocation_Received <- vapply(strsplit(coalesce(bt$Allocation, ""), ",\\s*"),
                                   function(a) collapse_u(unname(recv_map[a])), character(1))
  list(
    paths     = paths,
    extra_err = extra_err,
    extras    = c(`Backup inventory` = !is.null(bk), `Old final reports` = !is.null(old)),
    path     = live,
    mtime    = mtime,
    loaded   = Sys.time(),
    sheets   = sheets,
    bt       = bt,
    final_reports = afr,
    index    = build_lookup_index(bt),
    ol_sids  = unique(norm_sid(ol$Sample_ID[nz(ol$Sample_ID)])),
    awaiting = aw,
    overview = ov,
    recv_map = recv_map,
    unknown_runs = unknown_runs,
    bc_dates  = bc_dates,
    unk_dates = unk_dates,
    sample_runs = sample_runs,
    realloc_sids = realloc_sids,
    sample_total_runs = run_maps$total,
    sanger   = try_read(read_sanger(path)),
    orig     = ol,
    rlog_rows = rlog_raw_rows,
    req_issues = req_issues,
    req_rows   = length(unique(req$Input)),
    overview_auto = tryCatch(build_overview_auto(ol, afr, gsa, wgs, ov, recv_map), error = function(e)
      structure(tibble(), load_error = conditionMessage(e))),
    dropped_repeat = tryCatch(dropped_repeat(ol, recv_map), error = function(e)
      structure(tibble(), load_error = conditionMessage(e))),
    ol_fr_issues = tryCatch(ol_fr_check(ol, afr), error = function(e)
      structure(tibble(), load_error = conditionMessage(e))),
    rerun_id_issues = tryCatch(rerun_id_issues(rlog, bt), error = function(e)
      structure(tibble(), load_error = conditionMessage(e))),
    reruns   = reruns_tbl,
    requested = req,
    reextracted = rx_tbl,
    reextract = tryCatch(build_reextract(ol, bt, rlog, reruns_tbl, rx_tbl), error = function(e)
      structure(tibble(), load_error = conditionMessage(e)))
  )
}

# -----------------------------------------------------------------------------
# Table display
# -----------------------------------------------------------------------------

HTML_COLS <- c("APOE", "ABCA7", "GSA", "WGS", "Backup", "ID check")

# ID Finder columns that can be shown or hidden
AQ_COLS <- c("Input", "SampleID", "SampleID (no dash)", "AliquotID", "Allocation", "Barcode", "Notes",
             "Status (Original List)", "Matched on")

# Column groups the Sample lookup can show or hide (names as display_tbl makes them)
LOOKUP_GROUPS <- list(
  `Search (Input, Matched via)`  = c("Input", "Matched via"),
  `IDs (Barcode, SampleID, AliquotID)` = c("Barcode", "SampleID", "AliquotID"),
  `Allocation & status` = c("Allocation", "Allocation received", "Status",
                            "Same SampleID, other allocation", "Re-tubed (same allocation)", "Comments"),
  `Sanger (APOE, ABCA7)` = c("APOE", "ABCA7"),
  `GSA`    = c("GSA", "GSA handoff", "GSA back"),
  `WGS`    = c("WGS", "WGS type", "WGS handoff", "WGS back"),
  `Backup` = c("Backup", "Backup plate", "Backup pos", "Backup note"),
  `Checks (ID check, Original List match)` = c("ID check", "Original List match"),
  `Plate, position, substudy, sex` = c("Plate", "Position", "Substudy", "Sex")
)

display_tbl <- function(d, extra_first = NULL) {
  out <- d %>%
    transmute(
      Barcode, SampleID, AliquotID = Aliquot, Allocation,
      `Allocation received` = Allocation_Received,
      Status,
      `Same SampleID, other allocation` = Other_Alloc_Barcodes,
      `Re-tubed (same allocation)` = Retubed_With,
      Comments = OL_Comments,
      APOE  = pill(APOE, APOE_raw),
      ABCA7 = pill(ABCA7, ABCA7_raw),
      GSA   = pill(GSA_Stage),
      `GSA handoff` = GSA_Handoff,
      `GSA back`    = GSA_Back,
      WGS   = pill(WGS_Stage),
      `WGS type`    = WGS_Type,
      `WGS handoff` = WGS_Handoff,
      `WGS back`    = WGS_Back,
      Backup        = pill(Backup_Stage),
      `Backup plate` = Backup_Plate,
      `Backup pos`  = Backup_Pos,
      `Backup note` = Backup_Note,
      `ID check` = ifelse(is.na(ID_conflict), "",
                          paste(pill(rep("Mismatch", length(ID_conflict))),
                                htmltools::htmlEscape(coalesce(ID_conflict, "")))),
      Plate, Position, Substudy, Sex,
      `Original List match` = OL_match
    )
  if (!is.null(extra_first)) out <- bind_cols(extra_first, out)
  out
}

# How many columns to freeze on the left (navbar "Freeze" menu). Read inside a
# render, so changing the menu redraws the tables.
freeze_n <- function() {
  s <- shiny::getDefaultReactiveDomain()
  v <- if (is.null(s)) 0 else suppressWarnings(as.integer(s$input$freeze_cols))
  if (!length(v) || is.na(v)) 0L else max(0L, min(v, 6L))
}

dt <- function(df, html_cols = intersect(HTML_COLS, names(df)), page = 25, filter = "top") {
  nfix <- min(freeze_n(), ncol(df))
  exts <- c(if (ALLOW_EXPORT) "Buttons", if (nfix > 0) "FixedColumns")
  datatable(
    df,
    rownames   = FALSE,
    escape     = setdiff(names(df), html_cols),
    filter     = filter,
    extensions = if (length(exts)) exts else list(),
    class      = "compact stripe hover nowrap",
    options    = c(list(
      dom = if (ALLOW_EXPORT) "Blfrtip" else "lfrtip",
      # Copy without a title line, so it pastes straight into Excel (header + all rows)
      buttons = if (ALLOW_EXPORT) list(list(extend = "copy", title = NULL), "csv", "excel") else list(),
      # page = -1 shows every row on one page
      pageLength = page,
      lengthMenu = list(c(25, 50, 100, -1), c("25", "50", "100", "All")),
      scrollX = TRUE, scrollY = TABLE_HEIGHT, scrollCollapse = TRUE, autoWidth = FALSE,
      rowCallback = CHECK_JS
    ), if (nfix > 0) list(fixedColumns = list(leftColumns = nfix)))
  )
}

# Field/value view of one or more tracker rows (e.g. "38754" and "38754_1"
# both belong to allocation 38754 and are shown side by side).
kv_table <- function(rows) {
  if (!nrow(rows)) return(NULL)
  labels <- if ("Allocation" %in% names(rows)) coalesce(rows$Allocation, "") else rep("", nrow(rows))
  labels <- if (nrow(rows) == 1) "Value" else make.unique(ifelse(nzchar(labels), labels, "Row"), sep = " #")
  out <- data.frame(Field = names(rows), stringsAsFactors = FALSE, check.names = FALSE)
  for (i in seq_len(nrow(rows))) {
    out[[labels[i]]] <- vapply(rows[i, , drop = FALSE],
                               function(x) coalesce(as.character(x), ""), character(1), USE.NAMES = FALSE)
  }
  keep <- apply(out[, -1, drop = FALSE], 1, function(r) any(r != ""))
  out[keep, , drop = FALSE]
}

alloc_sort <- function(x) {
  x <- unique(x[nz(x)])
  num <- suppressWarnings(as.numeric(x))
  c(x[!is.na(num)][order(-num[!is.na(num)])], sort(x[is.na(num)]))
}

# -----------------------------------------------------------------------------
# UI
# -----------------------------------------------------------------------------

css <- "
.pill{display:inline-block;padding:.08rem .55rem;border-radius:999px;font-size:.76rem;
      font-weight:600;white-space:nowrap;line-height:1.5}
.pill-ok{background:#E3F2E8;color:#1E6B3A}
.pill-bad{background:#FBE4E4;color:#A12C2C}
.pill-warn{background:#FFF1D6;color:#8A5A00}
.pill-info{background:#E2EEF9;color:#1F5F8B}
.pill-flag{background:#F1E8FB;color:#6B3FA0}
.pill-drop{background:#F2ECEC;color:#7A5A5A}
.pill-muted{background:#EEF0F2;color:#5B6470}
.status-line{font-size:.8rem;color:#5B6470}
.notfound{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.85rem}
textarea#ids{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.85rem}
table.dataTable td{vertical-align:middle}
/* drag the bottom-right corner of any panel to make it taller or shorter */
.resizable{resize:vertical;overflow:auto;min-height:140px}
.resizable.bslib-full-screen{resize:none}
.path-pop{--bs-popover-max-width:480px}
.section-split{display:flex;align-items:center;gap:.75rem;margin:2rem 0 1rem;color:#7A4A00;font-weight:600;
  text-transform:uppercase;letter-spacing:.04em;font-size:.85rem}
.section-split::before,.section-split::after{content:'';flex:1;border-top:2px dashed #D9A441}
.realloc-card{border:2px solid #D9A441 !important}
.realloc-head{background:#FFF4DC !important;color:#7A4A00;font-weight:600;border-bottom:2px solid #D9A441}
table.dataTable th.col-divider, table.dataTable td.col-divider{border-left:3px solid #1F5F8B !important}
.freeze-pick{position:fixed;left:12px;bottom:10px;z-index:1050;background:var(--bs-body-bg);
  border:1px solid var(--bs-border-color);border-radius:6px;padding:3px;box-shadow:0 1px 4px rgba(0,0,0,.12)}
.freeze-pick .form-group{margin-bottom:0}
.freeze-pick select{font-size:.85rem;padding:.2rem .5rem}
table.dataTable .dtfc-fixed-left{background-color:var(--bs-body-bg) !important}
.dataTables_scrollBody::-webkit-scrollbar{height:14px;width:12px}
.dataTables_scrollBody::-webkit-scrollbar-thumb{background:#B8C2CC;border-radius:7px;border:3px solid transparent;background-clip:padding-box}
.dataTables_scrollBody::-webkit-scrollbar-thumb:hover{background:#8A97A5;background-clip:padding-box}
.dataTables_scrollBody{scrollbar-width:auto;scrollbar-color:#B8C2CC transparent}
"

# Every panel: expand button (bottom-right, on hover) + drag-to-resize height.
# Contents are NOT fill items: a table inside a fill container whose height is
# never set can collapse to zero height and look blank.
panel <- function(title = NULL, ...) {
  card(
    if (!is.null(title)) card_header(title),
    card_body(..., fillable = FALSE),
    full_screen = TRUE, fill = FALSE, class = "resizable"
  )
}

`%||%` <- function(a, b) if (is.null(a)) b else a

# -----------------------------------------------------------------------------
# Shared data cache
# One copy of the data serves every open browser. The workbooks are re-read only
# when one changes on disk; a failed read keeps the last good copy and is retried
# when the file changes again (or after a minute).
# -----------------------------------------------------------------------------

CACHE <- new.env(parent = emptyenv())
CACHE$paths        <- INITIAL_PATHS
CACHE$mtimes       <- NULL
CACHE$data         <- NULL
CACHE$err          <- NULL
CACHE$data_version <- 0L
CACHE$last_try     <- Sys.time()

path_mtimes <- function(p) {
  vapply(p, function(x) if (nzchar(x) && file.exists(x)) as.numeric(file.mtime(x)) else NA_real_,
         numeric(1))
}

cache_store <- function(res) {
  if (inherits(res, "error")) {
    CACHE$err <- conditionMessage(res)
    message(format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "  Lab Tracker: load failed: ", CACHE$err)
  } else {
    CACHE$data <- res
    CACHE$err  <- NULL
    CACHE$data_version <- CACHE$data_version + 1L
    if (length(res$extra_err)) {
      message(format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "  Lab Tracker: ",
              paste(sprintf("%s: %s", names(res$extra_err), unlist(res$extra_err)), collapse = " | "))
    }
  }
}

cache_refresh <- function(force = FALSE) {
  m <- path_mtimes(CACHE$paths)
  retry <- !is.null(CACHE$err) &&
    as.numeric(difftime(Sys.time(), CACHE$last_try, units = "secs")) > 60
  if (!force && !retry && identical(CACHE$mtimes, m)) return(invisible(FALSE))
  CACHE$mtimes   <- m
  CACHE$last_try <- Sys.time()
  cache_store(tryCatch(load_tracker(CACHE$paths), error = function(e) e))
  invisible(TRUE)
}

nav_items <- list(
  
  nav_panel(
    "Sample lookup",
    layout_sidebar(
      sidebar = sidebar(
        width = 320,
        textAreaInput("ids", "Barcodes", rows = 14,
                      placeholder = "Paste barcodes, one per line\n(SampleIDs and aliquot IDs also work)"),
        actionButton("go", "Look up", class = "btn-primary w-100"),
        helpText("Matching ignores case, spaces and hyphens. A SampleID or aliquot returns every barcode tied to it."),
        # collapsible so the list can be tucked away; click the header to open it
        accordion(
          id = "lookup_groups_acc", open = FALSE,
          accordion_panel(
            "Columns to show", value = "cols", icon = NULL,
            checkboxGroupInput("lookup_groups", NULL, choices = names(LOOKUP_GROUPS),
                               selected = names(LOOKUP_GROUPS)),
            div(class = "d-flex gap-2",
                actionLink("lookup_groups_all", "All"), span("\u00b7"),
                actionLink("lookup_groups_none", "None"))
          )
        )
      ),
      uiOutput("lookup_boxes"),
      panel("Results", DTOutput("lookup_tbl")),
      uiOutput("lookup_notfound")
    )
  ),
  
  nav_panel(
    "Allocations",
    layout_sidebar(
      sidebar = sidebar(
        width = 260,
        uiOutput("alloc_picker"),
        checkboxInput("alloc_show_all", "Show all rows on one page", TRUE),
        checkboxInput("alloc_hide_dropped", "Hide rows whose Status starts with \"Dropped\"", FALSE),
        helpText("\"D-ropped\" rows are left visible (not matched by the filter)."),
        radioButtons("alloc_layout", "Tracker panels",
                     c("Side by side" = "side", "Stacked" = "stacked", "Hide" = "hide"),
                     selected = "side")
      ),
      uiOutput("alloc_boxes"),
      uiOutput("alloc_panels"),
      panel("Barcodes in this allocation", DTOutput("alloc_tbl")),
      # clearly separate from the table above: a labelled divider, space, and
      # a coloured card
      div(class = "section-split", span("Samples with an existing call")),
      card(
        card_header(class = "realloc-head",
                    "Samples here that also have an APOE / ABCA7 call in another allocation"),
        card_body(div(class = "status-line mb-1", textOutput("alloc_realloc_caption", inline = TRUE)),
                  DTOutput("alloc_realloc"), fillable = FALSE),
        full_screen = TRUE, fill = FALSE, class = "resizable realloc-card"
      )
    )
  ),
  
  nav_panel(
    "Queues",
    navset_card_tab(
      full_screen = TRUE,
      # grow with the table instead of a fixed ~400px box that scrolls inside
      wrapper = function(...) card_body(..., fillable = FALSE),
      nav_panel("Awaiting GSA", DTOutput("q_awaiting")),
      nav_panel("At GSA (not back)", DTOutput("q_gsa")),
      nav_panel("At WGS in-house (not back)", DTOutput("q_wgs")),
      nav_panel(
        "Dropped 2+ times",
        div(class = "status-line mb-1", textOutput("droprep_caption", inline = TRUE)),
        p(class = "status-line",
          "Samples whose Status starts with \u201cDropped\u201d in 2 or more allocations of Original List, with each ",
          "drop's allocation, reason and AllocationReceived date, and the sample's other allocations."),
        DTOutput("droprep_tbl")
      )
    )
  ),
  
  nav_panel(
    "Sanger Hiatus",
    div(class = "status-line mb-2", textOutput("rerun_caption", inline = TRUE)),
    navset_card_tab(
      full_screen = TRUE,
      wrapper = function(...) card_body(..., fillable = FALSE),
      nav_panel("APOE", DTOutput("rerun_apoe")),
      nav_panel("ABCA7", DTOutput("rerun_abca7")),
      nav_panel(
        "Requested Reruns",
        div(class = "status-line mb-1", textOutput("reqlist_caption", inline = TRUE)),
        p(class = "status-line",
          "Everything in the Requested Reruns table (Sanger_Rerun_List tab), both genes. An open request stays on ",
          "the Sanger Hiatus list even after a Pass in Original List; once its Result says Pass or Done it comes off."),
        checkboxInput("req_open_only", "Show only open requests (hide Done / Pass)", value = FALSE),
        DTOutput("reqlist_tbl")
      ),
      nav_panel(
        "Re-extraction",
        div(class = "status-line mb-2", textOutput("reex_caption", inline = TRUE)),
        p(class = "status-line",
          "Still failing, and the newest allocation's tube has been run 2+ times after a reallocation ",
          "(the sample has 2+ allocations; tubes sent straight to USUHS / NCRAD / CU don't count). ",
          "Same as \u201cRe-extract\u201d in the Recommended column of the APOE / ABCA7 lists. ",
          "Once a sample is re-extracted (\u201cRe-extracted\u201d in its Status or Comments in Original List, or listed ",
          "in the Re-extracted Samples table) it leaves this tab, but stays on the APOE / ABCA7 lists with a green ",
          "\u2713 until it passes."),
        DTOutput("reex_tbl")
      ),
      nav_panel(
        "Sample lookup",
        layout_columns(
          col_widths = c(3, 9),
          div(
            textAreaInput("h_ids", "Barcodes / SampleIDs", rows = 10,
                          placeholder = "Paste barcodes, SampleIDs or aliquot IDs, one per line"),
            actionButton("h_go", "Look up", class = "btn-primary w-100"),
            helpText("A SampleID shows all its tubes, plus an UNKNOWN row for reruns logged without a barcode.")
          ),
          div(DTOutput("h_lookup_tbl"), uiOutput("h_notfound"))
        )
      )
    ),
    p(class = "status-line mt-2",
      "Listed: samples with a Sanger Fail for that gene and no Pass for it in any allocation, plus everything in the ",
      "Requested Reruns table (even after a Pass). ",
      "Left out: \u201cFail, Dropped\u201d (not repeated) and samples whose every row is Dropped. ",
      "Times Run = how many times the sample appears in Sanger_Rerun_List for that gene, at least 1 (blank Gene counts as APOE).")
  ),
  
  nav_panel(
    "ID check",
    navset_card_tab(
      full_screen = TRUE,
      wrapper = function(...) card_body(..., fillable = FALSE),
      nav_panel(
        "Sheets disagree on SampleID",
        p(class = "status-line",
          "The app does not pick a winner. Original List columns for these barcodes follow the ",
          "ALL_FinalReports SampleID, so treat their status and Sanger results as unverified until resolved."),
        DTOutput("idcheck_tbl")
      ),
      nav_panel(
        "Original List vs final reports",
        div(class = "status-line mb-1", textOutput("olfr_caption", inline = TRUE)),
        p(class = "status-line",
          "Per allocation (in both Original List and ALL_FinalReports): samples in the final report but not in ",
          "Original List, Original List samples that aren't Dropped but have no final report row, Dropped samples ",
          "that still have a tube, and samples listed more times than they have barcodes (repeats are fine when each ",
          "has its own barcode in the final report). Allocations with no final report yet are listed once."),
        DTOutput("olfr_tbl")
      ),
      nav_panel(
        "Sanger_Rerun_List mismatches",
        div(class = "status-line mb-1", textOutput("rerun_id_caption", inline = TRUE)),
        p(class = "status-line",
          "Rows where the barcode and SampleID don't match. Each barcode is one tube: on the Sanger Hiatus tab ",
          "these barcodes are listed under the sample that owns them, but Times Run still counts by the SampleID ",
          "written in the list, so fixing these rows fixes both."),
        DTOutput("rerun_id_tbl")
      ),
      nav_panel(
        "Requested Reruns",
        div(class = "status-line mb-1", textOutput("req_caption", inline = TRUE)),
        p(class = "status-line",
          "Rows in the Requested Reruns table (Sanger_Rerun_List tab) to check: Gene blank or not APOE / ABCA7 / Both ",
          "(these are listed for both genes), or a SampleID the app can't find anywhere. They still appear on the Sanger Hiatus list."),
        DTOutput("req_tbl")
      )
    )
  ),
  
  nav_panel(
    "Overview",
    navset_card_tab(
      full_screen = TRUE,
      wrapper = function(...) card_body(..., fillable = FALSE),
      nav_panel("Automatic",
                div(class = "status-line mb-1", textOutput("overview_auto_caption", inline = TRUE)),
                p(class = "status-line",
                  "Built by the app from Original List, ALL_FinalReports, GSA and WGS, so every allocation is ",
                  "included and counted the same way as the rest of the app (\u201c_#\u201d folded into the parent; ",
                  "Pass / Fail variants recognised; no call = Fail, not counting \u201cFail, Dropped\u201d, which is no ",
                  "longer part of the working allocation). Columns marked ",
                  "(manual) are copied from Overview_Tracker_2.0. At GSA / at WGS = Handed Back blank or No."),
                DTOutput("overview_auto")),
      nav_panel("Manual",
                div(class = "status-line mb-2", textOutput("overview_caption", inline = TRUE)),
                DTOutput("overview_full"))
    )
  ),
  # tools for preparing work, rather than looking things up
  nav_menu(
    "Lab Tools",
    nav_panel(
      "Submission prep",
      layout_sidebar(
        fillable = FALSE,
        sidebar = sidebar(
          width = 320, open = "always",
          textAreaInput("sp_codes", "Barcodes (one per line, in plate order)", rows = 12,
                        placeholder = "FA28478598\nFA28478386\n..."),
          textInput("sp_gsa", "GSA plate name", value = paste0("AD_GSAv3MDNeuro_", format(Sys.Date(), "%b%Y"))),
          textInput("sp_wgs", "WGS plate name", value = paste0("AD_HIHG_WGS_", toupper(format(Sys.Date(), "%b%Y")))),
          textInput("sp_letter", "First plate letter", value = "A", width = "120px"),
          textInput("sp_ceph", "CEPH well(s) on every plate", value = "H12"),
          textInput("sp_ceph_id", "CEPH Sample ID (GENOLIST)", value = "19-960561"),
          if (ALLOW_EXPORT) downloadButton("sp_download", "Download submission files (.xlsx)", class = "btn-primary w-100")
          else p(class = "status-line", "Downloads are turned off on this server (LAB_TRACKER_ALLOW_EXPORT).")
        ),
        div(class = "status-line mb-2", textOutput("sp_caption", inline = TRUE)),
        navset_card_tab(
          wrapper = function(...) card_body(..., fillable = FALSE),
          nav_panel("Check",        sp_part_ui("check"),  DTOutput("sp_check")),
          nav_panel("GSA manifest", sp_part_ui("gsa"),    DTOutput("sp_gsa_tbl")),
          nav_panel("WGS manifest", sp_part_ui("wgs"),    DTOutput("sp_wgs_tbl")),
          nav_panel("STAPLe",       sp_part_ui("staple"), DTOutput("sp_staple_tbl")),
          nav_panel("GENOLIST",     sp_part_ui("geno"),   DTOutput("sp_geno_tbl"))
        )
      )
    ),
    nav_panel(
      "PCR calculator",
      layout_sidebar(
        fillable = FALSE,
        sidebar = sidebar(
          width = 300, open = "always",
          numericInput("pcr_n", "Number of reactions (your extra already included)", value = 96, min = 1, step = 1),
          tags$strong("Per reaction (\u00b5L)"),
          lapply(seq_along(PCR_RECIPE), function(i)
            numericInput(paste0("pcr_v", i), names(PCR_RECIPE)[i], value = unname(PCR_RECIPE[i]), min = 0, step = 0.05)),
          numericInput("pcr_dna", "DNA (added per well)", value = PCR_DNA_UL, min = 0, step = 0.5),
          actionLink("pcr_reset", "Reset to the standard recipe")
        ),
        card(card_header("Master mix"), card_body(tableOutput("pcr_mix"), fillable = FALSE), fill = FALSE),
        card(card_header("Clean-up (ExoSAP-IT)"),
             card_body(div(class = "d-flex gap-3 mb-2",
                           numericInput("clean_prod", "PCR product per reaction (\u00b5L)", value = 5, min = 0, width = "220px"),
                           numericInput("clean_exo", "ExoSAP-IT per reaction (\u00b5L)", value = 2, min = 0, width = "220px")),
                       tableOutput("pcr_clean"), fillable = FALSE), fill = FALSE)
      )
    ),
    nav_panel(
      "Volume \u2194 concentration",
      layout_columns(
        col_widths = c(6, 6, 12), fill = FALSE,
        card(card_header("Amount = concentration \u00d7 volume"),
             card_body(fillable = FALSE,
                       radioButtons("vc_solve", "Solve for", inline = TRUE,
                                    choices = c("Volume" = "vol", "Concentration" = "conc", "Amount" = "amt")),
                       div(class = "d-flex gap-2",
                           numericInput("vc_amt", "Amount", value = 1, min = 0, width = "140px"),
                           selectInput("vc_amt_u", "\u00a0", c(names(MASS_U), names(MOL_U)), selected = "ug", width = "90px")),
                       div(class = "d-flex gap-2",
                           numericInput("vc_conc", "Concentration", value = 50, min = 0, width = "140px"),
                           selectInput("vc_conc_u", "\u00a0", c(names(CONC_U), names(MCONC_U)), selected = "ng/uL", width = "110px")),
                       div(class = "d-flex gap-2",
                           numericInput("vc_vol", "Volume", value = 20, min = 0, width = "140px"),
                           selectInput("vc_vol_u", "\u00a0", names(VOL_U), selected = "uL", width = "90px")),
                       p(class = "status-line", "The answer is given in the unit picked next to the box being solved."),
                       tags$hr(), uiOutput("vc_out"))),
        card(card_header("Dilution: C1 \u00d7 V1 = C2 \u00d7 V2"),
             card_body(fillable = FALSE,
                       radioButtons("dl_solve", "Solve for", inline = TRUE,
                                    choices = c("V1 (stock to take)" = "v1", "V2 (final volume)" = "v2",
                                                "C2 (final conc.)" = "c2", "C1 (stock conc.)" = "c1")),
                       div(class = "d-flex gap-2",
                           numericInput("dl_c1", "C1 stock", value = 100, min = 0, width = "120px"),
                           selectInput("dl_c1_u", "\u00a0", c(names(CONC_U), names(MCONC_U), "X"), selected = "ng/uL", width = "100px"),
                           numericInput("dl_v1", "V1", value = 10, min = 0, width = "110px"),
                           selectInput("dl_v1_u", "\u00a0", names(VOL_U), selected = "uL", width = "80px")),
                       div(class = "d-flex gap-2",
                           numericInput("dl_c2", "C2 final", value = 10, min = 0, width = "120px"),
                           selectInput("dl_c2_u", "\u00a0", c(names(CONC_U), names(MCONC_U), "X"), selected = "ng/uL", width = "100px"),
                           numericInput("dl_v2", "V2", value = 100, min = 0, width = "110px"),
                           selectInput("dl_v2_u", "\u00a0", names(VOL_U), selected = "uL", width = "80px")),
                       p(class = "status-line", "Mixing mass (ng/uL ...) with molar (nM, uM, mM) uses the molecular weight below; ",
                         "X only works with X."),
                       tags$hr(), uiOutput("dl_out"))),
        card(card_header("Molecular weight (only needed to go between mass and moles)"),
             card_body(fillable = FALSE,
                       div(class = "d-flex flex-wrap gap-3 align-items-end",
                           numericInput("mw_val", "MW (g/mol)", value = NA, min = 0, width = "160px"),
                           span(class = "status-line pb-2", "or work it out from length:"),
                           numericInput("mw_len", "Length", value = NA, min = 0, width = "110px"),
                           radioButtons("mw_kind", NULL, inline = TRUE,
                                        choices = c("nt (single-stranded, e.g. primer)" = "ss", "bp (double-stranded)" = "ds")),
                           uiOutput("mw_note"))))
      )
    ),
    nav_panel(
      "ID Finder",
      layout_sidebar(
        fillable = FALSE,
        sidebar = sidebar(
          width = 300, open = "always",
          textAreaInput("aq_codes", "AliquotIDs, SampleIDs or barcodes (one per line)", rows = 14,
                        placeholder = "202504587-47\n20-2416923\n202416923\nFA08209390"),
          accordion(
            id = "aq_cols_acc", open = FALSE,
            accordion_panel(
              "Columns to show", value = "cols", icon = NULL,
              checkboxGroupInput("aq_cols", NULL, choices = AQ_COLS,
                                 selected = setdiff(AQ_COLS, "Status (Original List)")),
              div(class = "d-flex gap-2",
                  actionLink("aq_cols_all", "All"), span("\u00b7"),
                  actionLink("aq_cols_none", "None"))
            )
          ),
          p(class = "status-line",
            "Like Finder_2: a SampleID gives every aliquot, allocation and barcode for it; an AliquotID or barcode gives that tube.")
        ),
        div(class = "status-line mb-2", textOutput("aq_caption", inline = TRUE)),
        DTOutput("aq_tbl")
      )
    )
  ),
  nav_panel("Original List", panel("Original List",
                                   div(class = "status-line mb-2", textOutput("orig_caption", inline = TRUE)),
                                   DTOutput("orig_tbl"))),
  # raw view of any sheet: desktop only
  if (!SERVER_MODE) nav_panel(
    "Sheets",
    p(class = "status-line mb-2", "Raw view of a sheet as Excel last saved it."),
    navset_card_tab(
      full_screen = TRUE,
      wrapper = function(...) card_body(..., fillable = FALSE),
      nav_panel("Main sheets",   uiOutput("sheet_picker"),       DTOutput("sheet_tbl")),
      nav_panel("Allocations",   uiOutput("alloc_sheet_picker"), DTOutput("alloc_sheet_tbl"))
    )
  ),
  
  nav_spacer(),
  nav_item(uiOutput("status_line")),
  
  # choosing files and forcing a reload: desktop only
  if (!SERVER_MODE) nav_item(
    popover(
      actionButton("file_btn", "File", icon = icon("folder-open"),
                   class = "btn-sm btn-outline-secondary me-1"),
      title = "Workbooks",
      id = "file_pop",
      placement = "bottom",
      options = list(customClass = "path-pop"),
      textInput("path", "Lab tracker", value = INITIAL_PATHS$tracker, width = "100%",
                placeholder = "C:/Users/you/OneDrive/Lab_Tracker_2026.xlsx"),
      textInput("path_backup", "Backup inventory database (optional)", value = INITIAL_PATHS$backup,
                width = "100%", placeholder = ".../BackupInventoryDatabase.xlsx"),
      textInput("path_old", "Old final reports (optional)", value = INITIAL_PATHS$old,
                width = "100%", placeholder = ".../OLD_FinalReports.xlsx"),
      actionButton("load_path", "Load", class = "btn-primary btn-sm"),
      helpText(class = "mt-2 mb-0",
               "Leave an optional box empty to skip that file. ",
               "Windows: Shift + right-click the file \u2192 Copy as path. ",
               "Mac: select it in Finder \u2192 Option-Cmd-C. Quotes are fine.")
    )
  ),
  if (!SERVER_MODE) nav_item(actionButton("reload", "Reload", class = "btn-sm btn-outline-secondary"))
)

ui <- do.call(page_navbar, c(
  list(
    title = "Lab Tracker",
    theme = bs_theme(version = 5, primary = "#1F5F8B",
                     base_font = font_collection("system-ui", "-apple-system", "Segoe UI", "Roboto", "sans-serif")),
    header = tagList(
      tags$head(tags$style(HTML(css))),
      # freeze the first columns of every table so they stay put when scrolling
      # sideways; fixed in the bottom-left corner of the screen
      div(class = "freeze-pick",
          selectInput("freeze_cols", NULL, width = "175px", selectize = FALSE,
                      choices = c("No frozen columns" = 0, "Freeze 1 column" = 1, "Freeze 2 columns" = 2,
                                  "Freeze 3 columns" = 3, "Freeze 4 columns" = 4, "Freeze 5 columns" = 5)))
    ),
    fillable = FALSE
  ),
  Filter(Negate(is.null), nav_items)
))


# -----------------------------------------------------------------------------
# Server
# -----------------------------------------------------------------------------

server <- function(input, output, session) {
  
  data_rv      <- reactiveVal(NULL)   # this browser's view of the shared data
  load_err     <- reactiveVal(NULL)
  seen_version <- reactiveVal(-1L)
  
  sync_from_cache <- function() {
    if (!identical(CACHE$data_version, isolate(seen_version()))) {
      seen_version(CACHE$data_version)
      data_rv(CACHE$data)
    }
    load_err(CACHE$err)
  }
  
  # Check the workbooks every few seconds; re-read only if one changed.
  observe({
    invalidateLater(POLL_MS)
    cache_refresh()
    sync_from_cache()
  })
  
  observeEvent(input$reload, {
    if (SERVER_MODE) return()
    cache_refresh(force = TRUE)
    sync_from_cache()
  })
  
  # Switch workbooks (desktop only). New paths are only adopted if the lab
  # tracker loads, so a typo can't leave the app watching nothing. The
  # SERVER_MODE check matters: hiding the button alone wouldn't stop a browser
  # from sending this input.
  observeEvent(input$load_path, {
    if (SERVER_MODE) return()
    p <- list(tracker = clean_path(input$path),
              backup  = clean_path(input$path_backup),
              old     = clean_path(input$path_old))
    labels <- c(tracker = "Lab tracker", backup = "Backup inventory", old = "Old final reports")
    if (!nzchar(p$tracker)) {
      showNotification("Paste the lab tracker's path first.", type = "warning", id = "path_msg")
      return()
    }
    for (n in names(p)) {
      if (!nzchar(p[[n]])) next
      if (!file.exists(p[[n]])) {
        showNotification(sprintf("%s: no file found at %s", labels[[n]], p[[n]]),
                         type = "error", duration = 8, id = "path_msg")
        return()
      }
      if (!grepl("\\.xls[xm]?$", p[[n]], ignore.case = TRUE)) {
        showNotification(sprintf("%s: that doesn't look like an Excel file (.xlsx).", labels[[n]]),
                         type = "error", duration = 8, id = "path_msg")
        return()
      }
    }
    m <- path_mtimes(p)
    res <- tryCatch(load_tracker(p), error = function(e) e)
    if (inherits(res, "error")) {
      showNotification(paste("Couldn't read the lab tracker:", conditionMessage(res)),
                       type = "error", duration = NULL, id = "path_msg")
      return()  # still showing the previous data, which is fine
    }
    CACHE$paths    <- p
    CACHE$mtimes   <- m
    CACHE$last_try <- Sys.time()
    cache_store(res)
    sync_from_cache()
    save_paths(lapply(p, function(x) if (nzchar(x)) normalizePath(x, winslash = "/", mustWork = FALSE) else ""))
    updateTextInput(session, "path", value = normalizePath(p$tracker, winslash = "/", mustWork = FALSE))
    toggle_popover("file_pop", show = FALSE)
    ee <- res$extra_err
    if (length(ee)) {
      showNotification(paste0("Loaded the lab tracker, but ",
                              paste(sprintf("%s: %s", names(ee), unlist(ee)), collapse = " | ")),
                       type = "warning", duration = NULL, id = "path_msg")
    } else {
      showNotification("Loaded.", type = "message", id = "path_msg")
    }
  })
  
  td <- reactive({
    msg <- if (SERVER_MODE) {
      "The lab tracker data isn't available right now. Please try again in a few minutes."
    } else {
      paste("No workbook loaded. Click \"File\" (top right) and paste the path to Lab_Tracker_2026.xlsx.",
            if (!is.null(load_err())) paste0("\n\nLast error: ", load_err()) else "")
    }
    validate(need(!is.null(data_rv()), msg))
    data_rv()
  })
  
  # In server mode the header never shows paths or raw errors (those go to the
  # R console of whoever is running the app).
  output$status_line <- renderUI({
    d <- data_rv()
    err <- load_err()
    txt <- if (is.null(d)) "Not loaded" else if (SERVER_MODE)
      sprintf("Read-only · updated %s", format(d$mtime, "%b %d %H:%M")) else
        sprintf("Read-only · %s · saved %s", basename(d$path), format(d$mtime, "%b %d %H:%M"))
    tip <- function(x) if (SERVER_MODE) NULL else x
    tagList(
      span(class = "status-line me-2", title = if (!is.null(d)) tip(d$path), txt),
      if (!is.null(d)) lapply(names(d$extras), function(n) {
        e <- d$extra_err[[n]]
        key <- if (n == "Backup inventory") "backup" else "old"
        if (!is.null(e)) span(class = "pill pill-bad me-1", title = tip(e), paste(n, "unavailable"))
        else if (isTRUE(d$extras[[n]])) span(class = "pill pill-ok me-1", title = tip(d$paths[[key]]), paste("+", n))
      }),
      if (!is.null(err) && !is.null(d)) span(class = "pill pill-warn", title = tip(err),
                                             "Showing the last good copy")
    )
  })
  
  # ---- dropdowns ----
  # Built together with the page (renderUI), so they also work behind the login,
  # where the page doesn't exist until someone signs in. Rebuilt when the data
  # changes, keeping the current choice if it still exists.
  keep_or_first <- function(sel, opts, fallback = opts[1]) {
    if (length(sel) == 1 && !is.na(sel) && sel %in% opts) sel else fallback
  }
  
  output$alloc_picker <- renderUI({
    d <- td()
    allocs <- alloc_sort(unlist(strsplit(coalesce(d$bt$Allocation, ""), ",\\s*")))
    selectizeInput("alloc", "Allocation", choices = allocs,
                   selected = keep_or_first(isolate(input$alloc), allocs))
  })
  
  output$sheet_picker <- renderUI({
    req(!SERVER_MODE)
    d <- td()
    sh <- intersect(MAIN_SHEETS, d$sheets)
    div(style = "max-width:320px",
        selectInput("sheet", "Sheet", choices = sh,
                    selected = keep_or_first(isolate(input$sheet), sh)))
  })
  
  output$alloc_sheet_picker <- renderUI({
    req(!SERVER_MODE)
    d <- td()
    sh <- alloc_sort(d$sheets[is_alloc_sheet(d$sheets)])
    validate(need(length(sh) > 0, "No allocation sheets found."))
    div(style = "max-width:320px",
        selectInput("alloc_sheet", "Allocation", choices = sh,
                    selected = keep_or_first(isolate(input$alloc_sheet), sh)))
  })
  
  # ---- Sample lookup ----
  query <- eventReactive(input$go, input$ids)
  
  lookup <- reactive({
    req(query())
    d <- td()
    toks <- unique(trimws(unlist(strsplit(query(), "[[:space:],;]+"))))
    toks <- toks[nz(toks)]
    validate(need(length(toks) > 0, "Paste at least one barcode."))
    # q2 also restores a SampleID's dropped leading zero ("20-908705")
    q <- tibble(Input = toks, q = norm_id(toks), q2 = norm_sid(toks), ord = seq_along(toks))
    hits <- bind_rows(
      q %>% inner_join(d$index, by = "q", relationship = "many-to-many"),
      q %>% filter(q2 != q) %>% select(-q) %>% rename(q = q2) %>%
        inner_join(d$index, by = "q", relationship = "many-to-many")
    ) %>%
      distinct(Input, key, .keep_all = TRUE) %>%
      inner_join(d$bt, by = "key") %>%
      arrange(ord, Barcode)
    # A searched SampleID also gets a row for every Original List allocation
    # that none of its barcodes covers (including samples with no barcode at
    # all): Original List columns filled, tube columns blank.
    no_bc <- q %>% filter(!(q %in% d$index$q | q2 %in% d$index$q))
    ol    <- d$orig %>% filter(nz(Sample_ID)) %>%
      mutate(q2 = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID")))
    covered <- hits %>% select(Input, Allocation) %>%
      separate_rows(Allocation, sep = ",\\s*") %>% filter(nz(Allocation)) %>%
      distinct(Input, a = Allocation)
    only_ol <- q %>% inner_join(ol, by = "q2", relationship = "many-to-many") %>%
      anti_join(covered, by = c("Input", "a"))
    if (nrow(only_ol)) {
      extra <- as.data.frame(d$bt)[rep(NA_integer_, nrow(only_ol)), , drop = FALSE]
      rownames(extra) <- NULL
      extra$Input         <- only_ol$Input
      extra$ord           <- only_ol$ord
      extra$via           <- "SampleID (Original List only)"
      extra$SampleID      <- only_ol$Sample_ID
      extra$Allocation    <- alloc_base(col_or_na(only_ol, "AllocationID"))
      extra$Allocation_Received <- coalesce(col_or_na(only_ol, "AllocationReceived"),
                                            unname(d$recv_map[extra$Allocation]))
      extra$Status        <- col_or_na(only_ol, "Status")
      extra$OL_Comments   <- col_or_na(only_ol, "Comments")
      extra$Reallocated   <- if_else(only_ol$q2 %in% d$realloc_sids, CHECK_YES, NA_character_)
      # dropped with no tube: say so in the Barcode column instead of leaving it blank
      # ("D-ropped" is intentionally not matched)
      extra$Barcode       <- ifelse(grepl("^\\s*Dropped", coalesce(extra$Status, "")), "N/A", NA_character_)
      extra$APOE_raw      <- col_or_na(only_ol, "APOESanger(Result)")
      extra$ABCA7_raw     <- col_or_na(only_ol, "ABCA7Sanger(Result)")
      extra$APOE          <- sanger_cat(extra$APOE_raw)
      extra$ABCA7         <- sanger_cat(extra$ABCA7_raw)
      extra$Substudy      <- col_or_na(only_ol, "Substudy")
      extra$Sex           <- col_or_na(only_ol, "Sex")
      extra$OL_match      <- "Original List only (no barcode)"
      extra$Other_Alloc_Barcodes <- other_alloc_label(only_ol$q2, extra$Allocation, d$bt, d$orig)
      hits <- bind_rows(hits, extra) %>% arrange(ord, Barcode)
    }
    missing <- no_bc %>% filter(!q2 %in% ol$q2) %>% mutate(Note = "Not found in any loaded file")
    list(hits = hits, missing = missing, n_in = length(toks))
  })
  
  output$lookup_boxes <- renderUI({
    req(query())
    l <- lookup()
    found_inputs <- length(unique(l$hits$Input))
    layout_columns(
      fill = FALSE,
      value_box("Matched", sprintf("%d / %d", found_inputs, l$n_in), theme = "primary"),
      value_box("Barcodes returned", sum(nz(l$hits$Barcode) & !l$hits$Barcode %in% c("N/A", "UNKNOWN"))),
      value_box("Not found", nrow(l$missing), theme = if (nrow(l$missing)) "warning" else NULL),
      value_box("SampleID mismatches", sum(!is.na(l$hits$ID_conflict)),
                theme = if (any(!is.na(l$hits$ID_conflict))) "danger" else NULL)
    )
  })
  
  observeEvent(input$lookup_groups_all,
               updateCheckboxGroupInput(session, "lookup_groups", selected = names(LOOKUP_GROUPS)))
  observeEvent(input$lookup_groups_none,
               updateCheckboxGroupInput(session, "lookup_groups", selected = character()))
  
  output$lookup_tbl <- renderDT({
    l <- lookup()
    validate(need(nrow(l$hits) > 0, "No matches."))
    validate(need(length(input$lookup_groups) > 0, "Pick at least one column group on the left."))
    tb <- display_tbl(l$hits, extra_first = tibble(Input = l$hits$Input, `Matched via` = l$hits$via))
    keep <- unlist(LOOKUP_GROUPS[input$lookup_groups], use.names = FALSE)
    # any column not in a group (added later) stays visible
    keep <- c(keep, setdiff(names(tb), unlist(LOOKUP_GROUPS, use.names = FALSE)))
    dt(tb[, intersect(names(tb), keep), drop = FALSE], filter = "none")
  }, server = FALSE)
  
  output$lookup_notfound <- renderUI({
    req(query())
    m <- lookup()$missing
    if (!nrow(m)) return(NULL)
    card(
      card_header(sprintf("Not found (%d)", nrow(m))),
      tags$table(class = "table table-sm notfound",
                 tags$tbody(lapply(seq_len(nrow(m)), function(i)
                   tags$tr(tags$td(m$Input[i]), tags$td(m$Note[i])))))
    )
  })
  
  # ---- Allocations ----
  alloc_rows <- reactive({
    req(input$alloc)
    bt <- td()$bt
    keep <- vapply(strsplit(coalesce(bt$Allocation, ""), ",\\s*"),
                   function(s) input$alloc %in% s, logical(1))
    out <- bt[keep, , drop = FALSE]
    # Original List samples of this allocation with no barcode here (e.g. dropped
    # before a final report): Original List columns filled, Barcode "N/A" if dropped
    ol <- td()$orig
    if (!is.null(ol) && nrow(ol)) {
      ola <- ol %>% filter(nz(Sample_ID)) %>%
        mutate(q2 = norm_sid(Sample_ID), a = alloc_base(col_or_na(., "AllocationID"))) %>%
        filter(a == input$alloc, !q2 %in% norm_sid(coalesce(out$SampleID, "")))
      if (nrow(ola)) {
        extra <- as.data.frame(bt)[rep(NA_integer_, nrow(ola)), , drop = FALSE]
        rownames(extra) <- NULL
        extra$SampleID      <- ola$Sample_ID
        extra$Allocation    <- ola$a
        extra$Allocation_Received <- coalesce(col_or_na(ola, "AllocationReceived"),
                                              unname(td()$recv_map[ola$a]))
        extra$Status        <- col_or_na(ola, "Status")
        extra$OL_Comments   <- col_or_na(ola, "Comments")
        # ("D-ropped" is intentionally not matched)
        extra$Barcode       <- ifelse(grepl("^\\s*Dropped", coalesce(extra$Status, "")), "N/A", NA_character_)
        extra$APOE_raw      <- col_or_na(ola, "APOESanger(Result)")
        extra$ABCA7_raw     <- col_or_na(ola, "ABCA7Sanger(Result)")
        extra$APOE          <- sanger_cat(extra$APOE_raw)
        extra$ABCA7         <- sanger_cat(extra$ABCA7_raw)
        extra$Substudy      <- col_or_na(ola, "Substudy")
        extra$Sex           <- col_or_na(ola, "Sex")
        extra$OL_match      <- "Original List only (no barcode)"
        extra$Other_Alloc_Barcodes <- other_alloc_label(ola$q2, extra$Allocation, bt, ol)
        out <- bind_rows(out, extra)
      }
    }
    if (isTRUE(input$alloc_hide_dropped)) out <- out %>% filter(!grepl("^Dropped", coalesce(Status, "")))
    out
  })
  
  output$alloc_boxes <- renderUI({
    a <- alloc_rows()
    layout_columns(
      fill = FALSE,
      value_box("Barcodes", sum(valid_bc(a$Barcode) & !a$Barcode %in% "N/A"), theme = "primary",
                if (any(is.na(a$key))) p(class = "mb-0", sprintf("+ %d with no barcode", sum(is.na(a$key))))),
      value_box("Sanger: both Pass", sum(a$APOE %in% "Pass" & a$ABCA7 %in% "Pass")),
      value_box("Sanger: Fail (to redo)", {
        n_fail <- sum(a$APOE %in% "Fail" | a$ABCA7 %in% "Fail")
        n_drop <- sum(!(a$APOE %in% "Fail" | a$ABCA7 %in% "Fail") &
                        (a$APOE %in% "Fail, dropped" | a$ABCA7 %in% "Fail, dropped"))
        if (n_drop) sprintf("%d (+%d dropped)", n_fail, n_drop) else as.character(n_fail)
      }),
      # sent = everything handed off; out = still at that location (not handed back)
      value_box("GSA sent / out",
                sprintf("%d / %d", sum(a$GSA_Stage %in% c("At GSA", "Returned")), sum(a$GSA_Stage %in% "At GSA"))),
      value_box("WGS sent / out",
                sprintf("%d / %d", sum(a$WGS_Stage %in% c("At WGS", "Returned")), sum(a$WGS_Stage %in% "At WGS")),
                # USUHS samples are never handed back, so they aren't "out"
                if (any(a$WGS_Stage %in% "Sent to USUHS"))
                  p(class = "mb-0", sprintf("+ %d sent to USUHS", sum(a$WGS_Stage %in% "Sent to USUHS")))),
      value_box("In backup", sum(a$Backup_Stage %in% "In backup"))
    )
  })
  
  output$alloc_panels <- renderUI({
    ov <- panel("Overview_Tracker_2.0", DTOutput("alloc_overview"))
    st <- panel("Sanger_Tracker", DTOutput("alloc_sanger"))
    switch(input$alloc_layout %||% "side",
           side    = layout_columns(col_widths = c(6, 6), ov, st),
           stacked = tagList(ov, st),
           hide    = NULL)
  })
  
  output$alloc_overview <- renderDT({
    ov <- td()$overview
    validate(need(nrow(ov) > 0, coalesce(attr(ov, "load_error"), "Overview_Tracker_2.0 not available")))
    row <- ov[alloc_base(coalesce(ov$Allocation, "")) == input$alloc, , drop = FALSE]
    validate(need(nrow(row) > 0, "No row for this allocation."))
    datatable(kv_table(row), rownames = FALSE, class = "compact",
              options = list(dom = "t", paging = FALSE, ordering = FALSE, rowCallback = CHECK_JS))
  })
  
  output$alloc_sanger <- renderDT({
    st <- td()$sanger
    validate(need(nrow(st) > 0, coalesce(attr(st, "load_error"), "Sanger_Tracker not available")))
    row <- st[alloc_base(coalesce(st$Allocation, "")) == input$alloc, , drop = FALSE]
    validate(need(nrow(row) > 0, "No row for this allocation."))
    # NEED TO REDO: counted from Original List with the Sanger Hiatus tab's rule
    # (replaces the sheet's link to Sanger_Hiatus.xlsx / NOCALLSbyAllocation)
    r <- td()$reruns
    if (nrow(r)) {
      in_alloc <- vapply(strsplit(coalesce(r$`Failed in allocations`, ""), ",\\s*"),
                         function(a) input$alloc %in% alloc_base(a), logical(1))
      for (g in c("APOE", "ABCA7")) {
        cn <- find_col(row, paste("NEED TO REDO", g))
        if (!is.na(cn)) row[[cn]] <- as.character(sum(in_alloc & r$Gene == g))
      }
    }
    datatable(kv_table(row), rownames = FALSE, class = "compact",
              options = list(dom = "t", paging = FALSE, ordering = FALSE, rowCallback = CHECK_JS))
  })
  
  output$alloc_tbl <- renderDT({
    a <- alloc_rows()
    validate(need(nrow(a) > 0, "No barcodes."))
    dt(display_tbl(a), page = if (isTRUE(input$alloc_show_all)) -1 else 25)
  }, server = FALSE)
  
  # Samples in this allocation that are in another allocation (reallocated
  # before or after) where they already have an APOE or ABCA7 call, with that
  # allocation's tube and its GSA / WGS status.
  alloc_realloc <- reactive({
    req(input$alloc)
    a <- alloc_rows()
    d <- td()
    if (!nrow(a)) return(tibble())
    here <- unique(norm_sid(a$SampleID[nz(a$SampleID)]))
    ol <- d$orig
    if (is.null(ol) || !nrow(ol) || !length(here)) return(tibble())
    other <- ol %>% filter(nz(Sample_ID)) %>%
      transmute(sid_key = norm_sid(Sample_ID), SampleID = Sample_ID,
                `Other allocation` = alloc_base(col_or_na(., "AllocationID")),
                `Other status` = col_or_na(., "Status"),
                APOE_raw = col_or_na(., "APOESanger(Result)"), ABCA7_raw = col_or_na(., "ABCA7Sanger(Result)")) %>%
      filter(sid_key %in% here, nz(`Other allocation`), `Other allocation` != input$alloc,
             # a real call (Pass / Fail) for at least one gene; "N/A" or blank for both is skipped
             sanger_cat(APOE_raw) %in% c("Pass", "Fail", "Fail, dropped", "Multiple") |
               sanger_cat(ABCA7_raw) %in% c("Pass", "Fail", "Fail, dropped", "Multiple"))
    if (!nrow(other)) return(tibble())
    tubes <- d$bt %>% filter(nz(SampleID)) %>%
      mutate(sid_key = norm_sid(SampleID), a1 = alloc_base(sub(",.*$", "", coalesce(Allocation, "")))) %>%
      group_by(sid_key, a1) %>%
      summarise(`Other barcode` = collapse_u(Barcode[valid_bc(Barcode)]),
                GSA = collapse_u(GSA_Stage), WGS = collapse_u(WGS_Stage),
                `WGS type` = collapse_u(WGS_Type), .groups = "drop")
    here_bc <- a %>% filter(nz(SampleID)) %>% mutate(sid_key = norm_sid(SampleID)) %>%
      group_by(sid_key) %>% summarise(`Barcode here` = collapse_u(Barcode[valid_bc(Barcode) & Barcode != "N/A"]),
                                      .groups = "drop")
    other %>%
      left_join(tubes, by = c("sid_key", "Other allocation" = "a1")) %>%
      left_join(here_bc, by = "sid_key") %>%
      arrange(SampleID, `Other allocation`) %>%
      transmute(SampleID, `Barcode here`, `Other allocation`, `Other barcode`, `Other status`,
                APOE = pill(sanger_cat(APOE_raw), APOE_raw), ABCA7 = pill(sanger_cat(ABCA7_raw), ABCA7_raw),
                GSA = pill(GSA), WGS = pill(WGS), `WGS type`)
  })
  output$alloc_realloc_caption <- renderText({
    x <- alloc_realloc()
    if (!nrow(x)) return("None of this allocation's samples has an APOE / ABCA7 call in another allocation.")
    sprintf("%d samples (%d rows) already have a call in another allocation.", n_distinct(x$SampleID), nrow(x))
  })
  output$alloc_realloc <- renderDT({
    x <- alloc_realloc()
    validate(need(nrow(x) > 0, "None found."))
    dt(x, html_cols = c("APOE", "ABCA7", "GSA", "WGS"), page = if (isTRUE(input$alloc_show_all)) -1 else 25)
  }, server = FALSE)
  
  # ---- Queues ----
  output$q_awaiting <- renderDT({
    aw <- td()$awaiting
    validate(need(nrow(aw) > 0, coalesce(attr(aw, "load_error"), "Nothing awaiting GSA.")))
    dt(aw, html_cols = character(), page = 100)
  }, server = FALSE)
  
  output$q_gsa <- renderDT({
    a <- td()$bt %>% filter(GSA_Stage == "At GSA")
    validate(need(nrow(a) > 0, "Nothing currently at GSA."))
    dt(display_tbl(a), page = 100)
  }, server = FALSE)
  
  output$q_wgs <- renderDT({
    a <- td()$bt %>% filter(WGS_Stage == "At WGS")
    validate(need(nrow(a) > 0, "Nothing currently out for in-house WGS."))
    dt(display_tbl(a), page = 100)
  }, server = FALSE)
  
  # ---- Sanger Hiatus (rerun list) ----
  output$rerun_caption <- renderText({
    r <- td()$reruns
    if (!nrow(r)) return(coalesce(attr(r, "load_error"), "No samples need a Sanger rerun."))
    sprintf("%d APOE and %d ABCA7 samples to rerun. %d have already been run more than once.",
            sum(r$Gene == "APOE"), sum(r$Gene == "ABCA7"), sum(r$`Times Run` > 1))
  })
  
  rerun_table <- function(gene) {
    r <- td()$reruns
    validate(need(nrow(r) > 0, coalesce(attr(r, "load_error"), "No samples need a Sanger rerun.")))
    r <- r %>% filter(Gene == gene) %>% select(-Gene, -`Failed in allocations`, -Barcodes)
    for (cc in c("Requested", "Requester", "Note")) if (cc %in% names(r) && all(is.na(r[[cc]]))) r[[cc]] <- NULL
    validate(need(nrow(r) > 0, sprintf("No %s samples need a rerun.", gene)))
    rx <- td()$reextract
    if (nrow(rx)) {
      rx <- rx %>% filter(Gene == gene) %>% select(sid_key, State, `Fails per barcode`, Recommended)
      r <- r %>% mutate(sid_key = norm_sid(SampleID)) %>% left_join(rx, by = "sid_key") %>%
        mutate(`Re-extract` = reex_mark(State)) %>%
        select(SampleID, `Times Run`, `Re-extract`, `Fails per barcode`, everything(), -sid_key, -State) %>%
        relocate(Recommended, .after = last_col())
    }
    # after Recommended: was it sent to WGS, and is it also failing the other gene
    other <- if (gene == "APOE") "ABCA7" else "APOE"
    k <- norm_sid(r$SampleID)
    w <- td()$bt %>% filter(nz(SampleID), WGS_Stage %in% c("At WGS", "Returned", "Sent to USUHS")) %>%
      mutate(sid_key = norm_sid(SampleID),
             lbl = paste0(ifelse(valid_bc(Barcode), paste0(Barcode, ": "), ""),
                          case_when(WGS_Stage == "At WGS"        ~ "at WGS",
                                    WGS_Stage == "Sent to USUHS" ~ "sent to USUHS",
                                    TRUE                         ~ paste("returned", coalesce(WGS_Back, ""))),
                          ifelse(nz(WGS_Type), paste0(" (", WGS_Type, ")"), ""))) %>%
      group_by(sid_key) %>% summarise(lbl = collapse_u(lbl, "; "), .groups = "drop")
    wmap <- setNames(w$lbl, w$sid_key)
    r$`Sent to WGS` <- ifelse(k %in% names(wmap), paste("Yes -", unname(wmap[k])), "No")
    on_other <- td()$reruns %>% filter(Gene == other) %>% transmute(sid_key = norm_sid(SampleID))
    rx_o <- td()$reextract
    rec_o <- if (nrow(rx_o)) rx_o %>% filter(Gene == other) %>% select(sid_key, Recommended) else
      tibble(sid_key = character(), Recommended = character())
    rmap <- setNames(rec_o$Recommended, rec_o$sid_key)
    r[[paste("Also failing", other)]] <- ifelse(k %in% on_other$sid_key,
                                                ifelse(k %in% names(rmap), paste0("Yes (", unname(rmap[k]), ")"), "Yes"),
                                                "No")
    r$`Sanger result`  <- result_pill(r$`Sanger result`)
    w <- dt(r, html_cols = c("Sanger result", "Re-extract"), page = 100)
    # a dividing line before "Sent to WGS" (the extra info after Recommended)
    j <- match("Sent to WGS", names(r))
    if (!is.na(j)) w$x$options$columnDefs <- c(w$x$options$columnDefs,
                                               list(list(targets = j - 1, className = "col-divider")))
    w
  }
  output$rerun_apoe  <- renderDT(rerun_table("APOE"),  server = FALSE)
  
  # ---- Sanger Hiatus: Requested Reruns only ----
  req_list <- reactive({
    q <- td()$requested
    if (is.null(q) || !nrow(q)) return(tibble())
    r <- td()$reruns
    on <- if (nrow(r)) r %>% transmute(Gene, sid_key = norm_sid(SampleID), `Times Run`, `Rerun dates`, Allocations,
                                       Barcodes, Location, on = TRUE)
    else tibble(Gene = character(), sid_key = character(), `Times Run` = integer(), `Rerun dates` = character(),
                Allocations = character(), Barcodes = character(), Location = character(),
                on = logical())
    q %>%
      transmute(Gene = g, sid_key, SampleID, Requester, Note, Resubmitted,
                Result = coalesce(Result, NA_character_)) %>%
      left_join(on, by = c("Gene", "sid_key")) %>%
      mutate(Status = case_when(coalesce(sanger_cat(Result) == "Pass", FALSE) ~ "Done (passed)",
                                req_closed(Result)                           ~ "Done",
                                coalesce(on, FALSE)                          ~ "Open (on Sanger Hiatus)",
                                TRUE                                         ~ "Open")) %>%
      arrange(Status != "Open (on Sanger Hiatus)", Gene, SampleID)
  })
  output$reqlist_caption <- renderText({
    r <- req_list()
    if (!nrow(r)) return("No Requested Reruns.")
    sprintf("%d requested (APOE %d, ABCA7 %d): %d done, %d still open.", nrow(r),
            sum(r$Gene == "APOE"), sum(r$Gene == "ABCA7"),
            sum(startsWith(r$Status, "Done")), sum(!startsWith(r$Status, "Done")))
  })
  output$reqlist_tbl <- renderDT({
    r <- req_list()
    validate(need(nrow(r) > 0, "No Requested Reruns."))
    if (isTRUE(input$req_open_only)) r <- r %>% filter(!startsWith(Status, "Done"))
    validate(need(nrow(r) > 0, "No open requests: every request is Done or passed."))
    r <- r %>% transmute(Gene, SampleID, Status, Result, `Resubmission date` = Resubmitted, Requester, Note,
                         `Times Run`, `Rerun dates`, Allocations, Barcodes, Location)
    r$Result <- result_pill(r$Result)
    dt(r, html_cols = "Result", page = 100)
  }, server = FALSE)
  output$rerun_abca7 <- renderDT(rerun_table("ABCA7"), server = FALSE)
  
  # ---- Sanger Hiatus: re-extraction ----
  reex_rows <- reactive({
    rx <- td()$reextract
    validate(need(is.null(attr(rx, "load_error")), paste("Couldn't check re-extraction:", attr(rx, "load_error"))))
    if (!nrow(rx)) return(rx)
    r <- td()$reruns
    # re-extracted samples are done here: they leave this tab, but stay on the
    # APOE / ABCA7 lists (marked Re-extracted) until the new extraction passes
    rx %>% filter(State %in% "Meets criteria") %>%
      left_join(r %>% transmute(Gene, sid_key = norm_sid(SampleID), `Times Run`, Allocations, Barcodes, Location),
                by = c("Gene", "sid_key")) %>%
      arrange(State == "Re-extracted", Gene, SampleID)
  })
  output$reex_caption <- renderText({
    x <- reex_rows()
    n_done <- sum(td()$reextract$State %in% "Re-extracted")
    if (!nrow(x)) return(sprintf("No samples currently meet the re-extraction criteria (%d already re-extracted).", n_done))
    sprintf("%d meet the criteria (APOE %d, ABCA7 %d). %d already re-extracted (not shown; still on the APOE / ABCA7 lists).",
            nrow(x), sum(x$Gene == "APOE"), sum(x$Gene == "ABCA7"), n_done)
  })
  output$reex_tbl <- renderDT({
    x <- reex_rows()
    validate(need(nrow(x) > 0, "No samples currently meet the re-extraction criteria."))
    x <- x %>% transmute(Gene, SampleID,
                         `Re-extract` = reex_mark(State),
                         `Times Run`, Allocations, `Fails per barcode`, Location)
    dt(x, html_cols = "Re-extract", page = 100)
  }, server = FALSE)
  
  # ---- Sanger Hiatus: sample lookup ----
  h_query <- eventReactive(input$h_go, input$h_ids)
  
  h_lookup <- reactive({
    req(h_query())
    d <- td()
    toks <- unique(trimws(unlist(strsplit(h_query(), "[[:space:],;]+"))))
    toks <- toks[nz(toks)]
    validate(need(length(toks) > 0, "Paste at least one barcode or SampleID."))
    q <- tibble(Input = toks, q = norm_id(toks), q2 = norm_sid(toks), ord = seq_along(toks))
    hits <- bind_rows(
      q %>% inner_join(d$index, by = "q", relationship = "many-to-many"),
      q %>% filter(q2 != q) %>% select(-q) %>% rename(q = q2) %>%
        inner_join(d$index, by = "q", relationship = "many-to-many")
    ) %>%
      distinct(Input, key, .keep_all = TRUE) %>%
      inner_join(d$bt, by = "key")
    
    tubes <- hits %>% transmute(
      Input, ord, sid = norm_sid(SampleID), SampleID, Barcode, AliquotID = Aliquot, Allocation,
      Status, Reallocated, APOE = pill(APOE, APOE_raw), ABCA7 = pill(ABCA7, ABCA7_raw),
      `Barcode runs (APOE / ABCA7)` = Reruns,
      `APOE rerun dates`  = unname(d$bc_dates$APOE[key]),
      `ABCA7 rerun dates` = unname(d$bc_dates$ABCA7[key]),
      Location = case_when(
        !is.na(Sent_To)               ~ NEVER_RECEIVED(Sent_To),
        GSA_Stage == "At GSA"         ~ "At GSA",
        WGS_Stage == "At WGS"         ~ "At WGS",
        WGS_Stage == "Sent to USUHS"  ~ "USUHS",
        GSA_Stage == "Awaiting GSA"   ~ "Awaiting GSA",
        Backup_Stage %in% "In backup" ~ paste("Backup", Backup_Plate, Backup_Pos),
        TRUE                          ~ NA_character_))
    
    # reruns logged with a SampleID but no barcode (only for SampleID searches)
    unk <- q %>% inner_join(d$unknown_runs, by = c("q2" = "sid_key")) %>%
      transmute(Input, ord, sid = q2, SampleID, Barcode = "UNKNOWN",
                # if_else keeps the column character even when there are no rows
                Reallocated = if_else(q2 %in% d$realloc_sids, CHECK_YES, NA_character_),
                `Barcode runs (APOE / ABCA7)` = runs,
                `APOE rerun dates`  = unname(d$unk_dates$APOE[q2]),
                `ABCA7 rerun dates` = unname(d$unk_dates$ABCA7[q2]))
    
    on_list <- d$reruns
    hiatus <- function(sid) {
      if (!nrow(on_list)) return(rep(NA_character_, length(sid)))
      k <- norm_sid(on_list$SampleID)
      vapply(sid, function(x) collapse_u(on_list$Gene[k %in% x]), character(1), USE.NAMES = FALSE)
    }
    rx <- d$reextract
    reex <- function(sid) {
      if (!nrow(rx)) return(rep(NA_character_, length(sid)))
      rx2 <- rx %>% filter(!is.na(State))
      vapply(sid, function(x) {
        h <- rx2[rx2$sid_key %in% x, ]
        if (!nrow(h)) return(NA_character_)
        paste(paste(h$Gene, reex_mark(h$State)), collapse = ", ")
      }, character(1), USE.NAMES = FALSE)
    }
    # on the Hiatus list (e.g. via Requested Reruns) but no barcode anywhere
    on_sids <- if (nrow(on_list)) norm_sid(on_list$SampleID) else character()
    # Original List allocations of the searched sample that no tube covers
    # (e.g. dropped before a final report): one row each, Barcode "N/A" if dropped
    olq <- if (!is.null(d$orig) && nrow(d$orig)) d$orig %>% filter(nz(Sample_ID)) %>%
      transmute(q2 = norm_sid(Sample_ID), SampleID = Sample_ID,
                Allocation = alloc_base(col_or_na(., "AllocationID")), Status = col_or_na(., "Status"),
                APOE_raw = col_or_na(., "APOESanger(Result)"), ABCA7_raw = col_or_na(., "ABCA7Sanger(Result)")) %>%
      filter(nz(Allocation))
    else tibble(q2 = character(), SampleID = character(), Allocation = character(), Status = character(),
                APOE_raw = character(), ABCA7_raw = character())
    covered <- tubes %>% transmute(sid, a = alloc_base(sub(",.*$", "", coalesce(Allocation, "")))) %>% distinct()
    olrows <- q %>% inner_join(olq, by = "q2", relationship = "many-to-many") %>%
      anti_join(covered, by = c("q2" = "sid", "Allocation" = "a")) %>%
      distinct(Input, Allocation, .keep_all = TRUE) %>%
      transmute(Input, ord, sid = q2, SampleID,
                Barcode = ifelse(grepl("^\\s*Dropped", coalesce(Status, "")), "N/A", "None found"),
                Allocation, Status,
                APOE = pill(sanger_cat(APOE_raw), APOE_raw), ABCA7 = pill(sanger_cat(ABCA7_raw), ABCA7_raw))
    # on the Hiatus list (e.g. via Requested Reruns) but no barcode or Original List row
    nobc <- q %>% filter(q2 %in% on_sids, !q2 %in% c(tubes$sid, unk$sid, olrows$sid)) %>%
      distinct(Input, .keep_all = TRUE) %>%
      transmute(Input, ord, sid = q2,
                SampleID = paste0(substr(q2, 1, 2), "-", substr(q2, 3, nchar(q2))),
                Barcode = "None found")
    out <- bind_rows(tubes, unk, olrows, nobc) %>%
      mutate(`Times in list (APOE / ABCA7)` = coalesce(unname(d$sample_runs[sid]), "0 / 0"),
             `Total runs (APOE / ABCA7)`   = coalesce(unname(d$sample_total_runs[sid]), "0 / 0"),
             `On Sanger Hiatus` = hiatus(sid),
             `Re-extract` = reex(sid)) %>%
      arrange(ord, SampleID, Barcode %in% c("UNKNOWN", "N/A", "None found"), Barcode) %>%
      select(Input, SampleID, Barcode, AliquotID, Allocation, Status, APOE, ABCA7,
             `Barcode runs (APOE / ABCA7)`, `APOE rerun dates`, `ABCA7 rerun dates`,
             `Times in list (APOE / ABCA7)`, `Total runs (APOE / ABCA7)`,
             `On Sanger Hiatus`, `Re-extract`, Location)
    missing <- q %>% filter(!(q %in% d$index$q | q2 %in% d$index$q | q2 %in% d$unknown_runs$sid_key |
                                q2 %in% on_sids | q2 %in% olq$q2))
    list(out = out, missing = missing)
  })
  
  output$h_lookup_tbl <- renderDT({
    l <- h_lookup()
    validate(need(nrow(l$out) > 0, "No matches."))
    dt(l$out, html_cols = c("APOE", "ABCA7", "Re-extract"), filter = "none", page = 100)
  }, server = FALSE)
  
  output$h_notfound <- renderUI({
    req(h_query())
    m <- h_lookup()$missing
    if (!nrow(m)) return(NULL)
    p(class = "status-line mt-2", sprintf("Not found (%d): %s", nrow(m), paste(m$Input, collapse = ", ")))
  })
  
  # ---- ID check ----
  output$droprep_caption <- renderText({
    x <- td()$dropped_repeat
    if (!is.null(attr(x, "load_error"))) return(paste("Couldn't check:", attr(x, "load_error")))
    if (!nrow(x)) return("No sample has been dropped in 2 or more allocations.")
    sprintf("%d samples dropped 2+ times (%d dropped 3+ times).", nrow(x), sum(x$`Times dropped` >= 3))
  })
  output$droprep_tbl <- renderDT({
    x <- td()$dropped_repeat
    validate(need(nrow(x) > 0, "No sample has been dropped in 2 or more allocations."))
    x$APOE  <- result_pill(x$APOE)
    x$ABCA7 <- result_pill(x$ABCA7)
    dt(x, html_cols = c("APOE", "ABCA7"), page = 100)
  }, server = FALSE)
  
  output$olfr_caption <- renderText({
    x <- td()$ol_fr_issues
    if (!is.null(attr(x, "load_error"))) return(paste("Couldn't compare:", attr(x, "load_error")))
    if (!nrow(x)) return("Original List and the final reports match: no missing or extra samples.")
    sprintf("%d flagged: %d missing from Original List, %d missing from the final reports, %d other.",
            nrow(x), sum(grepl("missing from Original List", x$Problem)),
            sum(grepl("missing from the final report", x$Problem)),
            sum(!grepl("missing from", x$Problem)))
  })
  output$olfr_tbl <- renderDT({
    x <- td()$ol_fr_issues
    validate(need(nrow(x) > 0, "No problems: every allocation's samples match between Original List and the final reports."))
    dt(x, html_cols = character(), page = 100)
  }, server = FALSE)
  
  output$rerun_id_caption <- renderText({
    d <- td()
    n_log <- if (is.null(d$rlog_rows)) 0 else d$rlog_rows
    sprintf("%d rows read from Sanger_Rerun_List; %d flagged.", n_log, nrow(d$rerun_id_issues))
  })
  
  output$req_caption <- renderText({
    d <- td()
    if (!is.null(attr(d$req_issues, "load_error"))) return(paste("Couldn't read Requested Reruns:", attr(d$req_issues, "load_error")))
    if (!d$req_rows) return("No Requested Reruns table found in Sanger_Rerun_List, or it is empty.")
    sprintf("%d samples in Requested Reruns; %d flagged.", d$req_rows, nrow(d$req_issues))
  })
  output$req_tbl <- renderDT({
    x <- td()$req_issues
    validate(need(nrow(x) > 0, "No problems found in Requested Reruns."))
    dt(x, html_cols = character(), page = 100)
  }, server = FALSE)
  
  output$rerun_id_tbl <- renderDT({
    x <- td()$rerun_id_issues
    validate(need(is.null(attr(x, "load_error")),
                  paste("Couldn't check Sanger_Rerun_List:", attr(x, "load_error"))))
    validate(need(nrow(x) > 0, "No problems: every barcode in Sanger_Rerun_List matches its SampleID."))
    dt(x, html_cols = character(), page = 100)
  }, server = FALSE)
  
  output$idcheck_tbl <- renderDT({
    a <- td()$bt %>% filter(!is.na(ID_conflict))
    validate(need(nrow(a) > 0, "No conflicts — every barcode maps to one SampleID across sheets."))
    out <- a %>% transmute(Barcode, Allocation,
                           `FinalReports SampleID` = afr_sid, `Old final reports SampleID` = old_sid,
                           `GSA SampleID` = gsa_sid, `WGS SampleID` = wgs_sid,
                           `Awaiting SampleID` = aw_sid, `Backup SampleID` = bk_sid,
                           `Backup plate` = Backup_Plate, `Backup pos` = Backup_Pos,
                           `Aliquot (FinalReports)` = coalesce(afr_aliquot, old_aliquot),
                           `Aliquot (Backup)` = bk_aliquot)
    dt(out, html_cols = character())
  }, server = FALSE)
  
  # ---- Overview / Original List / Sheets ----
  output$overview_caption <- renderText({
    ov <- td()$overview
    if (!nrow(ov)) return(paste("Nothing read from Overview_Tracker_2.0.", coalesce(attr(ov, "load_error"), "")))
    sprintf("%d allocations \u00d7 %d columns read from Overview_Tracker_2.0", nrow(ov), ncol(ov))
  })
  
  # ---- PCR calculator ----
  observeEvent(input$pcr_reset, {
    for (i in seq_along(PCR_RECIPE)) updateNumericInput(session, paste0("pcr_v", i), value = unname(PCR_RECIPE[i]))
    updateNumericInput(session, "pcr_dna", value = PCR_DNA_UL)
  })
  fmt_ul <- function(x) ifelse(is.na(x), "", formatC(x, format = "f", digits = 2, big.mark = ",", drop0trailing = TRUE))
  output$pcr_mix <- renderTable({
    n <- suppressWarnings(as.numeric(input$pcr_n)); if (is.na(n) || n < 0) n <- 0
    v <- vapply(seq_along(PCR_RECIPE), function(i) {
      x <- suppressWarnings(as.numeric(input[[paste0("pcr_v", i)]])); if (is.null(x) || is.na(x)) 0 else x }, numeric(1))
    dna <- suppressWarnings(as.numeric(input$pcr_dna)); if (is.na(dna)) dna <- 0
    tot <- v * n
    data.frame(
      Component = c(names(PCR_RECIPE), "DNA", "Total"),
      `1x (uL)` = fmt_ul(c(v, dna, sum(v) + dna)),
      `Master mix for N (uL)` = c(fmt_ul(tot), "added per well", fmt_ul(sum(tot))),
      `Half (uL)` = c(fmt_ul(tot / 2), "", fmt_ul(sum(tot) / 2)),
      `Quarter (uL)` = c(fmt_ul(tot / 4), "", fmt_ul(sum(tot) / 4)),
      check.names = FALSE)
  }, striped = TRUE, spacing = "s", align = "lrrrr")
  output$pcr_clean <- renderTable({
    n <- suppressWarnings(as.numeric(input$pcr_n)); if (is.na(n) || n < 0) n <- 0
    p <- suppressWarnings(as.numeric(input$clean_prod)); if (is.na(p)) p <- 0
    e <- suppressWarnings(as.numeric(input$clean_exo));  if (is.na(e)) e <- 0
    data.frame(Component = c("PCR Product", "ExoSAP-IT", "Solution total"),
               `1x (uL)` = fmt_ul(c(p, e, p + e)),
               `For N reactions (uL)` = fmt_ul(c(p, e, p + e) * n), check.names = FALSE)
  }, striped = TRUE, spacing = "s", align = "lrr")
  
  # ---- Volume <-> concentration ----
  num_in <- function(x) { v <- suppressWarnings(as.numeric(x)); if (!length(v) || is.na(v)) NA_real_ else v }
  f3 <- function(x) formatC(x, format = "g", digits = 4, big.mark = ",")
  # molecular weight: typed, or from length (average: ssDNA nt x 308.97 + 18.02; dsDNA bp x 617.96 + 36.04)
  mw <- reactive({
    v <- num_in(input$mw_val)
    if (!is.na(v) && v > 0) return(v)
    L <- num_in(input$mw_len)
    if (is.na(L) || L <= 0) return(NA_real_)
    if (identical(input$mw_kind, "ds")) L * 617.96 + 36.04 else L * 308.97 + 18.02
  })
  output$mw_note <- renderUI({
    m <- mw()
    if (is.na(m)) return(span(class = "status-line pb-2", "Not set."))
    span(class = "pb-2", HTML(sprintf("Using <b>%s g/mol</b>%s", f3(m),
                                      if (is.na(num_in(input$mw_val))) " (average for that length; use the exact MW from the oligo sheet if you have it)" else "")))
  })
  output$vc_out <- renderUI({
    ua <- input$vc_amt_u; uc <- input$vc_conc_u; uv <- input$vc_vol_u
    used <- c(amt = ua, conc = uc)[setdiff(c("amt", "conc"), input$vc_solve)]
    need_mw <- length(unique(is_molar(c(ua, uc)))) > 1
    m <- mw()
    if (need_mw && is.na(m))
      return(HTML("<div class='text-danger'>Mass and moles are mixed: enter the molecular weight (or length) below.</div>"))
    if (is.na(m)) m <- 1000          # all mass or all moles: the MW cancels out
    amt  <- to_ng(num_in(input$vc_amt), ua, m)      # ng
    conc <- to_ngul(num_in(input$vc_conc), uc, m)   # ng/uL
    vol  <- num_in(input$vc_vol) * VOL_U[[uv]]      # uL
    res <- switch(input$vc_solve,
                  vol  = if (isTRUE(conc > 0)) sprintf("Volume = <b>%s %s</b>", f3(amt / conc / VOL_U[[uv]]), uv)
                  else "Enter a concentration above 0.",
                  conc = if (isTRUE(vol > 0)) sprintf("Concentration = <b>%s %s</b>", f3(from_ngul(amt / vol, uc, m)), uc)
                  else "Enter a volume above 0.",
                  amt  = sprintf("Amount = <b>%s %s</b>", f3(from_ng(conc * vol, ua, m)), ua))
    HTML(paste0("<div style='font-size:1.15rem'>", res, "</div>"))
  })
  output$dl_out <- renderUI({
    u1 <- input$dl_c1_u; u2 <- input$dl_c2_u
    if ((u1 == "X") != (u2 == "X"))
      return(HTML("<div class='text-danger'>X can only be diluted to X.</div>"))
    m <- mw()
    if (u1 != "X" && is_molar(u1) != is_molar(u2) && is.na(m))
      return(HTML("<div class='text-danger'>Mass and moles are mixed: enter the molecular weight (or length) below.</div>"))
    if (is.na(m)) m <- 1000
    cv <- function(v, u) if (u == "X") v else to_ngul(v, u, m)
    cb <- function(x, u) if (u == "X") x else from_ngul(x, u, m)
    c1 <- cv(num_in(input$dl_c1), u1); c2 <- cv(num_in(input$dl_c2), u2)
    fv1 <- VOL_U[[input$dl_v1_u]]; fv2 <- VOL_U[[input$dl_v2_u]]
    v1 <- num_in(input$dl_v1) * fv1; v2 <- num_in(input$dl_v2) * fv2
    out <- switch(input$dl_solve,
                  v1 = if (isTRUE(c1 > 0)) { x <- c2 * v2 / c1
                  sprintf("Take <b>%s %s</b> of stock + <b>%s %s</b> water (to %s %s)",
                          f3(x / fv1), input$dl_v1_u, f3((v2 - x) / fv2), input$dl_v2_u, f3(v2 / fv2), input$dl_v2_u) }
                  else "Enter C1 above 0.",
                  v2 = if (isTRUE(c2 > 0)) { x <- c1 * v1 / c2
                  sprintf("Final volume <b>%s %s</b>: add <b>%s %s</b> water to the %s %s of stock",
                          f3(x / fv2), input$dl_v2_u, f3((x - v1) / fv2), input$dl_v2_u, f3(v1 / fv1), input$dl_v1_u) }
                  else "Enter C2 above 0.",
                  c2 = if (isTRUE(v2 > 0)) sprintf("Final concentration <b>%s %s</b>", f3(cb(c1 * v1 / v2, u2)), u2) else "Enter V2 above 0.",
                  c1 = if (isTRUE(v1 > 0)) sprintf("Stock concentration <b>%s %s</b>", f3(cb(c2 * v2 / v1, u1)), u1) else "Enter V1 above 0.")
    warn <- if (input$dl_solve == "v1" && isTRUE(c1 > 0) && isTRUE(c2 > c1))
      "<div class='text-danger'>C2 is higher than C1: you can't dilute up.</div>" else ""
    HTML(paste0("<div style='font-size:1.15rem'>", out, "</div>", warn))
  })
  
  # ---- ID Finder ----
  # Finder_2 in the app: rows come from ALL_FinalReports (SampleID, child_aliquot,
  # order_id, barcode_2d, notes). Matches an AliquotID, a barcode, or a SampleID
  # (with or without the dash). An aliquot that isn't in the final reports falls
  # back to its worked-out SampleID (202504587-47 -> 20-2504587).
  aq <- reactive({
    x <- trimws(unlist(strsplit(coalesce(input$aq_codes, ""), "[[:space:],;]+")))
    x <- x[nzchar(x)]
    if (!length(x)) return(NULL)
    d <- td()
    chr <- function(v) { v <- as.character(v); v[!nz(v)] <- NA_character_; v }
    fr <- d$final_reports
    fr <- tibble(SampleID   = chr(col_or_na(fr, "SampleID")),
                 AliquotID  = chr(col_or_na(fr, "child_aliquot")),
                 Allocation = chr(col_or_na(fr, "order_id")),
                 Barcode    = chr(col_or_na(fr, "barcode_2d")),
                 Notes      = chr(col_or_na(fr, "notes"))) %>%
      filter(nz(SampleID) | nz(AliquotID) | nz(Barcode)) %>%
      mutate(k_aq = norm_id(coalesce(AliquotID, "")), k_bc = norm_id(coalesce(Barcode, "")),
             k_sid = norm_sid(coalesce(SampleID, "")))
    q <- tibble(.ord = seq_along(x), Input = x, k = norm_id(x), k2 = norm_sid(x),
                k3 = norm_sid(coalesce(aliquot_to_sid(x), "")))
    m1 <- q %>% inner_join(fr, by = c(k = "k_aq"), relationship = "many-to-many") %>% mutate(`Matched on` = "AliquotID")
    m2 <- q %>% inner_join(fr, by = c(k = "k_bc"), relationship = "many-to-many") %>% mutate(`Matched on` = "Barcode")
    m3 <- q %>% inner_join(fr, by = c(k2 = "k_sid"), relationship = "many-to-many") %>% mutate(`Matched on` = "SampleID")
    cols <- c(".ord", "Input", "SampleID", "AliquotID", "Allocation", "Barcode", "Notes", "Matched on")
    hit <- bind_rows(m1[, cols], m2[, cols], m3[, cols]) %>% distinct(.ord, AliquotID, Barcode, Allocation, .keep_all = TRUE)
    rest <- q %>% filter(!.ord %in% hit$.ord)
    m4 <- rest %>% filter(nzchar(k3)) %>% inner_join(fr, by = c(k3 = "k_sid"), relationship = "many-to-many") %>%
      mutate(`Matched on` = "SampleID worked out from the aliquot")
    hit <- bind_rows(hit, m4[, cols])
    # SampleIDs on the Original List with no tube in the final reports for that allocation
    ol <- d$orig %>% filter(nz(Sample_ID)) %>%
      transmute(k_sid = norm_sid(Sample_ID), SampleID = chr(Sample_ID),
                Allocation = chr(col_or_na(., "AllocationID")), Status = chr(col_or_na(., "Status")))
    # an entry that is itself a SampleID uses it; otherwise the SampleID worked out from the aliquot
    want <- q %>% mutate(ks = ifelse(k2 %in% ol$k_sid | !nzchar(k3), k2, k3))
    seen <- hit %>% mutate(ks = norm_sid(coalesce(SampleID, "")), a = alloc_base(coalesce(Allocation, ""))) %>%
      distinct(ks, a)
    ol_only <- want %>% filter(!.ord %in% m1$.ord, !.ord %in% m2$.ord) %>%
      inner_join(ol, by = c(ks = "k_sid"), relationship = "many-to-many") %>%
      mutate(a = alloc_base(coalesce(Allocation, ""))) %>%
      anti_join(seen, by = c("ks", "a")) %>%
      transmute(.ord, Input, SampleID, AliquotID = NA_character_, Allocation, Barcode = NA_character_,
                Notes = NA_character_, `Matched on` = "Original List only (no tube in final reports)")
    hit <- bind_rows(hit, ol_only)
    none <- q %>% filter(!.ord %in% hit$.ord) %>%
      # as.character(): ifelse() on zero rows returns logical, which bind_rows can't mix with text
      transmute(.ord, Input,
                SampleID = as.character(ifelse(grepl("-", Input), aliquot_to_sid(Input), NA_character_)),
                AliquotID = NA_character_,
                Allocation = NA_character_, Barcode = NA_character_, Notes = NA_character_,
                `Matched on` = as.character(ifelse(!is.na(SampleID), "Not found (SampleID worked out from the aliquot)",
                                                   "Not found")))
    out <- bind_rows(hit, none) %>%
      mutate(`SampleID (no dash)` = ifelse(nz(SampleID), gsub("-", "", SampleID), NA_character_),
             n_a = suppressWarnings(as.numeric(alloc_base(coalesce(Allocation, ""))))) %>%
      arrange(.ord, n_a, Allocation, Barcode) %>% select(-n_a)
    if ("Status (Original List)" %in% input$aq_cols) {
      st <- ol %>% mutate(a = alloc_base(coalesce(Allocation, ""))) %>%
        group_by(k_sid, a) %>% summarise(Status = collapse_u(Status, "; "), .groups = "drop")
      out <- out %>% mutate(k_sid = norm_sid(coalesce(SampleID, "")), a = alloc_base(coalesce(Allocation, ""))) %>%
        left_join(st, by = c("k_sid", "a")) %>% select(-k_sid, -a) %>%
        rename(`Status (Original List)` = Status)
    }
    out %>% select(.ord, Input, SampleID, `SampleID (no dash)`, AliquotID, Allocation, Barcode, Notes,
                   any_of("Status (Original List)"), `Matched on`)
  })
  output$aq_caption <- renderText({
    x <- aq()
    if (is.null(x)) return("Paste AliquotIDs, SampleIDs or barcodes on the left.")
    nf <- n_distinct(x$.ord[startsWith(x$`Matched on`, "Not found")])
    sprintf("%d rows for %d entries%s.", nrow(x), n_distinct(x$.ord),
            if (nf) sprintf("; %d not found in the final reports or Original List", nf) else "")
  })
  observeEvent(input$aq_cols_all,  updateCheckboxGroupInput(session, "aq_cols", selected = AQ_COLS))
  observeEvent(input$aq_cols_none, updateCheckboxGroupInput(session, "aq_cols", selected = character()))
  output$aq_tbl <- renderDT({
    x <- aq(); validate(need(!is.null(x), "Paste AliquotIDs, SampleIDs or barcodes on the left."))
    validate(need(length(input$aq_cols) > 0, "Pick at least one column on the left."))
    dt(x[, intersect(AQ_COLS[AQ_COLS %in% input$aq_cols], names(x)), drop = FALSE],
       html_cols = character(), page = 100)
  }, server = FALSE)
  
  # ---- Submission prep ----
  sp <- reactive({
    codes <- unlist(strsplit(coalesce(input$sp_codes, ""), "[[:space:],;]+"))
    prep_submission(codes, td(), input$sp_gsa, input$sp_wgs, input$sp_letter, input$sp_ceph, input$sp_ceph_id)
  })
  output$sp_caption <- renderText({
    x <- sp()
    if (is.null(x)) return("Paste barcodes on the left to build the plates.")
    sprintf("%d tubes on %d plate(s)%s.", x$n, x$plates,
            if (x$warn) sprintf(" \u2014 %d with a warning (see Check)", x$warn) else ", no warnings")
  })
  sp_tbl <- function(part) renderDT({
    x <- sp(); validate(need(!is.null(x), "Paste barcodes on the left."))
    dt(x[[part]], html_cols = character(), page = 100)
  }, server = FALSE)
  output$sp_check     <- sp_tbl("check")
  output$sp_gsa_tbl   <- sp_tbl("gsa")
  output$sp_wgs_tbl   <- sp_tbl("wgs")
  output$sp_staple_tbl <- sp_tbl("staple")
  output$sp_geno_tbl  <- sp_tbl("geno")
  sp_names <- c(check = "Check", gsa = "GSA_Manifest", wgs = "WGS_Manifest",
                staple = "STAPLe", geno = "GENOLIST")
  for (part in names(sp_names)) local({
    pp <- part
    output[[paste0("sp_csv_", pp)]] <- downloadHandler(
      filename = function() paste0(sp_names[[pp]], "_", gsub("[^A-Za-z0-9_-]", "_", coalesce(input$sp_gsa, "plates")), ".csv"),
      content = function(file) {
        x <- sp(); validate(need(!is.null(x), "Paste barcodes first."))
        utils::write.csv(x[[pp]], file, row.names = FALSE, na = "")
      })
  })
  output$sp_download <- downloadHandler(
    filename = function() paste0("Submission_", gsub("[^A-Za-z0-9_-]", "_", coalesce(input$sp_gsa, "plates")), ".xlsx"),
    content = function(file) {
      x <- sp()
      validate(need(!is.null(x), "Paste barcodes first."))
      if (!requireNamespace("writexl", quietly = TRUE))
        stop("Install the writexl package first: install.packages(\"writexl\")")
      writexl::write_xlsx(list(GSA_Manifest_OUTPUT = x$gsa, WGS_Manifest_OUTPUT = x$wgs,
                               `GSA.WGS_STAPLe_OUTPUT` = x$staple, GENOLIST = x$geno, Check = x$check), file)
    }
  )
  
  output$overview_auto_caption <- renderText({
    x <- td()$overview_auto
    if (!is.null(attr(x, "load_error"))) return(paste("Couldn't build the overview:", attr(x, "load_error")))
    if (!nrow(x)) return("No allocations found.")
    sprintf("%d allocations, newest received first.", nrow(x))
  })
  output$overview_auto <- renderDT({
    x <- td()$overview_auto
    validate(need(nrow(x) > 0, "No allocations found."))
    dt(x, html_cols = character(), page = 100)
  }, server = FALSE)
  
  output$overview_full <- renderDT({
    ov <- td()$overview
    validate(need(nrow(ov) > 0, coalesce(attr(ov, "load_error"), "Overview_Tracker_2.0 not available")))
    dt(ov, html_cols = character(), page = 50)
  }, server = FALSE)
  
  output$orig_caption <- renderText({
    ol <- td()$orig
    sprintf("%d rows \u00d7 %d columns read from Original List. Use the boxes under each header to filter.",
            nrow(ol), ncol(ol))
  })
  
  output$orig_tbl <- renderDT({
    ol <- td()$orig
    validate(need(nrow(ol) > 0, "Original List is empty."))
    res_cols <- c(find_col(ol, "APOESanger(Result)"), find_col(ol, "ABCA7Sanger(Result)"))
    res_cols <- res_cols[!is.na(res_cols)]
    for (cn in res_cols) ol[[cn]] <- result_pill(ol[[cn]])
    dt(ol, html_cols = res_cols, page = 100)
  }, server = FALSE)
  
  # Raw sheet view (desktop only). Only sheets from the two lists can be opened.
  raw_sheet_dt <- function(sheet) {
    d <- td()
    snap <- snapshot_workbook(d$path)
    on.exit(unlink(snap), add = TRUE)
    m <- read_raw(snap, sheet)
    validate(need(nrow(m) > 0, "Sheet is empty."))
    hr <- which(rowSums(!is.na(m)) >= 2)[1]
    if (is.na(hr)) hr <- 1
    tbl <- drop_gsa_result(build_table(m, unlist(m[hr, ], use.names = FALSE), hr + 1, stop_at_blank = FALSE))
    datatable(tbl, rownames = FALSE, filter = "top", class = "compact stripe hover nowrap",
              options = list(pageLength = 25, scrollX = TRUE, scrollY = TABLE_HEIGHT,
                             scrollCollapse = TRUE, rowCallback = CHECK_JS,
                             lengthMenu = list(c(25, 50, 100, -1), c("25", "50", "100", "All"))))
  }
  
  output$sheet_tbl <- renderDT({
    req(!SERVER_MODE, input$sheet)
    req(input$sheet %in% intersect(MAIN_SHEETS, td()$sheets))
    raw_sheet_dt(input$sheet)
  })
  
  output$alloc_sheet_tbl <- renderDT({
    req(!SERVER_MODE, input$alloc_sheet)
    req(input$alloc_sheet %in% td()$sheets, is_alloc_sheet(input$alloc_sheet))
    raw_sheet_dt(input$alloc_sheet)
  })
}

# Server mode: read the workbooks at startup so the first visitor doesn't wait.
if (SERVER_MODE) cache_refresh(force = TRUE)

if (USE_LOGIN) {
  app_server <- function(input, output, session) {
    shinymanager::secure_server(check_credentials = shinymanager::check_credentials(CREDENTIALS))
    server(input, output, session)
  }
  shinyApp(shinymanager::secure_app(ui, enable_admin = FALSE), app_server)
} else {
  shinyApp(ui, server)
}