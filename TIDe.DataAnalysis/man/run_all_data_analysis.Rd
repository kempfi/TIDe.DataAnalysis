% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/run_all_data_analysis.R
\name{run_all_data_analysis}
\alias{run_all_data_analysis}
\title{run_all_data_analysis}
\usage{
run_all_data_analysis(
  excel_file,
  excel_sheet_name,
  excel_nrows = 100,
  excel_ncols = 180,
  activate_arrival_time_detailed_analysis = FALSE,
  activate_arrival_time_analysis_plots = TRUE
)
}
\arguments{
\item{excel_file}{Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"}

\item{excel_sheet_name}{Name of Excel Sheet}

\item{excel_nrows}{Number of rows to be copied out, after header (1:(excel_nrows+1)). Default: 100.}

\item{excel_ncols}{Number of columns to be copied out (1:excel_ncols). Default: 180.}

\item{activate_arrival_time_detailed_analysis}{Default \code{FALSE}. If false, then the detailed arrival time
analysis with gaussian fits is skipped. If true it is executed.}

\item{activate_arrival_time_analysis_plots}{Default \code{TRUE}. If true, the plots showing the arrival time data
analysis results with gaussian fits are automatically created. If false this step is skipped.}
}
\value{
Executed for its side effects.
}
\description{
This function runs a full data analysis over different measurements. It reads an Excel sheet and saves
it into a datframe. Afterwards it extracts measurement information from the file name, then a general data
analysis is performed and in the end the dark rate is also analyzed.
}
\details{
Note that the function only analyses "activated" rows. The user can activate a row by putting 1 into the respective
data analysis field (for example column C (Filename Read in), P (Data Analysis), CB (DR Data Analysis)). If a value of 0
is entered said calculation is skipped. This allows the user to save calculation time and not rerun results which have
already been calculated.

The user can also select the data to be filtered for unwanted noise by selecting 1 in column BU (or for dark rate column DW).
In the column directly after the user can input a value of x, corresponding to which ionizations shall be ignored for the
calculations. For example a value of x = 5 will mean that if per alpha particle 5 ionizations are detected it will not be
filtered out, but if there are 6 ionizations per alpha this will be interpreted as noise and be taken out of the calculation.
The number of ionizations counted, which get saved back into the excel, are not filtered. Hence the user can still directly check
how often such individual noises or unwanted signals happened.

Excel sheet must contain a sheet named "Info", where in cell B1 the
  GEM type of the measurements is given and in cell B2 th apth to the
  folder with measurement data is given. In the general measurement data folder,
  subfolders for each GEM should be present with the GEM name as folder name.

  Folder structure:
  \itemize{
  \item{.../Measurement_Data}
    \itemize{
    \item{em1mm}
    \item{em1p5mm}
    \item{em2mm}
    \item{keramik1p5mm}
  }}

  If alph_start and alph_end in the Excel file are not yet defined (TBD),
  the default values cover the entire range of the measurement (1:End).
}
\examples{
\dontrun{
detach("package:TIDe.DataAnalysis", unload=TRUE)
library("TIDe.DataAnalysis")

run_all_data_analysis(excel_file = "Example.xlsx",
                      excel_sheet_name = "M1",
                      excel_nrows = 100,
                      excel_ncols = 180)
}
}
