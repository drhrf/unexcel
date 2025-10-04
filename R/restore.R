#' Revert Excel date serials to intended day.month numerics
#'
#' Many spreadsheets auto-convert entries like '30.3' into dates. After import,
#' those values arrive as Excel date serials (integers). This function detects
#' such serials and reconstructs the intended 'day.month' decimals while
#' leaving other entries intact. Both 1900 and 1904 systems are supported.
#'
#' @param x A vector (numeric, integer, character, or Date).
#' @param low_serial Lower bound for plausible serials (inclusive).
#' @param high_serial Upper bound for plausible serials (inclusive).
#' @param year_window Integer vector of years that, when resolved, will be
#'   considered valid to revert. This guards against accidental matches.
#' @param origin_mode One of "auto", "1900", or "1904". In "1900" mode the
#'   origin is "1899-12-30" (Excel’s 1900 system with the leap-year quirk
#'   compensated). In "1904" mode the origin is "1904-01-01". In "auto" mode,
#'   the origin yielding dates with median proximity to a reference date is
#'   chosen; the reference can be controlled via \code{ref_date} for tests.
#' @param ref_date Reference date for origin selection when origin_mode="auto".
#'   Defaults to Sys.Date(); set to a fixed Date in tests for determinism.
#'
#' @return Returns a numeric vector when restoration is unambiguous; otherwise 
#' character vector, preserving both restored and untouched values.
#' @examples
#' restore_day_month(c(45812, 12.5, 44730), origin_mode = "1900")
#' restore_day_month(c(45812, 44730), origin_mode = "auto")
#' @export
restore_day_month <- function(
  x,
  low_serial = 20000,                 # ~1954-07-14 (1900 system)
  high_serial = 65000,                # ~2078-02-19
  year_window = 1990:2035,
  origin_mode = c("auto","1900","1904"),
  ref_date = Sys.Date()
) {
  origin_mode <- match.arg(origin_mode)
  origin_1900 <- "1899-12-30"
  origin_1904 <- "1904-01-01"

  # Handle Date directly for speed and clarity
  if (inherits(x, "Date")) {
    yy <- as.integer(format(x, "%Y"))
    keep <- !is.na(yy) & yy %in% year_window
    out <- as.character(x)
    if (any(keep)) {
      day <- sub("^0", "", format(x[keep], "%d"))
      mon <- sub("^0", "", format(x[keep], "%m"))
      out[keep] <- paste0(day, ".", mon)
    }
    out_num <- suppressWarnings(as.numeric(out))
    return(if (sum(!is.na(out_num)) >= sum(!is.na(x))) out_num else out)
  }

  xx <- suppressWarnings(as.numeric(x))
  is_num <- !is.na(xx) & is.finite(xx)
  is_intish <- is_num & abs(xx - round(xx)) < 1e-8
  in_range <- is_intish & xx >= low_serial & xx <= high_serial
  if (!any(in_range)) {
    out_num <- suppressWarnings(as.numeric(x))
    return(if (any(!is.na(out_num))) out_num else x)
  }

  choose_origin <- function() {
    if (origin_mode == "1900") return(origin_1900)
    if (origin_mode == "1904") return(origin_1904)
    d1900 <- as.Date(xx[in_range], origin = origin_1900)
    d1904 <- as.Date(xx[in_range], origin = origin_1904)
    m1 <- stats::median(abs(as.numeric(d1900 - ref_date)), na.rm = TRUE)
    m2 <- stats::median(abs(as.numeric(d1904 - ref_date)), na.rm = TRUE)
    if (m1 <= m2) origin_1900 else origin_1904
  }

  origin <- choose_origin()
  d <- as.Date(xx[in_range], origin = origin)
  yy <- as.integer(format(d, "%Y"))
  ok <- !is.na(yy) & yy %in% year_window

  out <- as.character(x)
  if (any(ok)) {
    d_ok <- d[ok]
    day <- sub("^0", "", format(d_ok, "%d"))
    mon <- sub("^0", "", format(d_ok, "%m"))
    repl <- paste0(day, ".", mon)
    out[which(in_range)[ok]] <- repl
  }

  out_num <- suppressWarnings(as.numeric(out))
  if (sum(!is.na(out_num)) >= sum(is_num)) out_num else out
}

#' Fix likely-serial columns in a data frame
#'
#' Applies \code{restore_day_month()} only to columns that look dominated by
#' Excel serials, controlled by a minimum fraction threshold.
#'
#' @param df A data frame.
#' @param pmin Minimum fraction of in-range integers to flag a column.
#' @inheritParams restore_day_month
#' @return The data frame with corrected columns where applicable.
#' @examples
#' df <- data.frame(a = c(45812, 44730), b = c(1.2, 3.4))
#' fix_serial_columns(df)
#' @export
fix_serial_columns <- function(
  df, pmin = 0.8,
  low_serial = 20000, high_serial = 65000,
  year_window = 1990:2035,
  origin_mode = "auto",
  ref_date = Sys.Date()
) {
  is_serial_col <- vapply(df, function(col) {
    num <- suppressWarnings(as.numeric(col))
    ok <- !is.na(num) & is.finite(num)
    if (!any(ok)) return(FALSE)
    intish <- abs(num[ok] - round(num[ok])) < 1e-8
    inrng <- num[ok] >= low_serial & num[ok] <= high_serial
    mean(intish & inrng) >= pmin
  }, logical(1))
  if (!any(is_serial_col)) return(df)
  df[is_serial_col] <- lapply(
    df[is_serial_col],
    restore_day_month,
    low_serial = low_serial, high_serial = high_serial,
    year_window = year_window, origin_mode = origin_mode, ref_date = ref_date
  )
  df
}