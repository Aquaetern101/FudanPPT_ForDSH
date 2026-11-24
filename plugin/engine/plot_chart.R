# Base R chart renderer (ASCII only; labels come from the CSV in UTF-8).
# usage: Rscript plot_chart.R <kind> <csv> <out.png> <palette> <w_in> <h_in> <dpi>
# csv: first column = label, remaining columns = series (header = series name)

args <- commandArgs(trailingOnly = TRUE)
kind  <- args[1]
incsv <- args[2]
out   <- args[3]
pal   <- strsplit(args[4], ",")[[1]]
w     <- as.numeric(args[5])
h     <- as.numeric(args[6])
dpi   <- as.numeric(args[7])

d <- read.csv(incsv, check.names = FALSE, fileEncoding = "UTF-8")
lab <- d[[1]]
mat <- as.matrix(d[, -1, drop = FALSE])
if (length(pal) < ncol(mat)) pal <- rep(pal, length.out = ncol(mat))
pal <- pal[seq_len(ncol(mat))]

png(out, width = w, height = h, units = "in", res = dpi, type = "cairo",
    family = "Microsoft YaHei", bg = "white")
par(mar = c(3.0, 4.2, 1.6, 1.0), las = 1, col.axis = "#333333",
    col.lab = "#333333", cex.axis = 1.05, cex.lab = 1.15, bty = "n")

if (kind == "bar") {
  bp <- barplot(t(mat), beside = TRUE, col = pal, border = NA,
                names.arg = lab, ylim = c(0, max(mat, na.rm = TRUE) * 1.18),
                ylab = "", xlab = "")
  axis(2, col = "#BBBBBB", col.ticks = "#BBBBBB")
  if (ncol(mat) > 1) {
    legend("topright", legend = colnames(mat), fill = pal, border = NA,
           bty = "n", cex = 0.95)
  }
} else if (kind == "line") {
  matplot(seq_len(nrow(mat)), mat, type = "o", pch = 16, lty = 1, lwd = 2,
          col = pal, xaxt = "n", ylab = "", xlab = "", ylim = c(0, max(mat, na.rm = TRUE) * 1.15))
  axis(1, at = seq_len(nrow(mat)), labels = lab, col = "#BBBBBB", col.ticks = "#BBBBBB")
  axis(2, col = "#BBBBBB", col.ticks = "#BBBBBB")
  if (ncol(mat) > 1) {
    legend("topright", legend = colnames(mat), col = pal, lty = 1, lwd = 2,
           bty = "n", cex = 0.95)
  }
} else if (kind == "pie") {
  par(mar = c(1.2, 1.2, 1.6, 1.2))
  pie(mat[, 1], labels = lab, col = pal, border = "white", cex = 1.05,
      init.angle = 90)
} else {
  stop(paste("unsupported kind:", kind))
}

dev.off()
cat("ok\n")
