#' @title add_gauss_fit_info_to_xslx
#' @description This function opens a certain sheet of an Excel file (.xlsx) and copies it out as a dataframe. It analyzes
#' the column "Fit Gauss Arrival Time", if this is activated (1), instead od deactivated (0), gaussian fit calculations are
#' performed. The stepsize is read out from the Excel Sheet, column "Fit Stepsize", if none is given a stepsize of 1000 is assumed.
#' The code then continues to calculate the arrival time data using \code{\link{get_data_arrival_time}}. As a default the last 10'000
#' ions are analyzed. If less data is available the code will analyze fewer curves and save the number of curves used into the excel file.
#'  In the end the fit parameters obtained with \code{\link{create_gaussian_fits}} and results are saved back into the xlsx file.
#'
#' @details Excel sheet must contain a sheet named "Info", where in cell B1 the
#'   GEM type of the measurements is given and in cell B2 th apth to the
#'   folder with measurement data is given. In the general measurement data folder,
#'   subfolders for each GEM should be present with the GEM name as folder name.
#'
#'   Folder structure:
#'   \itemize{
#'   \item{.../Measurement_Data}
#'     \itemize{
#'     \item{em1mm}
#'     \item{em1p5mm}
#'     \item{em2mm}
#'     \item{keramik1p5mm}
#'   }}
#'
#'   If alph_start and alph_end in the Excel file are not yet defined (TBD),
#'   the default values cover the entire range of the measurement (1:End).
#'
#' @param excel_file Excel file name
#' @param excel_sheet_name Name of Excel Sheet
#' @param excel_nrows Number of rows to be copied out, after header (1:(excel_nrows+1)). Default: 100.
#' @param excel_ncols Number of columns to be copied out (1:excel_ncols). Default: 180.
#' @param select_last_10000 If \code{TRUE} only the last 10000 Ions are selected for the Output. If \code{FALSE}
#'   all arrival times are exported.
#' @param activate_Plots if TRUE, plots are created, if FALSE, they are skipped.
#' @param alpha_packet_size Number of alpha particles grouped into one data point, default 100.
#' @param activate_save_back_to_xlsx Default \code{TRUE}. If True, then the dataframe with the gauss fits is saved
#' back to excel. If false this step is skipped.
#'
#' @return Dataframe with gauss fit information added to it.
#' @export
#' @importFrom("graphics", "plot", "hist")
#' @importFrom("stats", "sd", "lm")
#' @importFrom openxlsx loadWorkbook
#' @importFrom openxlsx read.xlsx
#' @importFrom openxlsx saveWorkbook
#' @importFrom openxlsx writeData
#' @examples \dontrun{
#' add_gauss_fit_info_to_xslx(excel_file = "em2mm_Data_Analysis.xlsx",
#'                            gem_type = "em2mm",
#'                            excel_sheet_name = "M2",
#'                            excel_nrows = 100,
#'                            activate_Plots = FALSE,
#'                            select_last_10000 = TRUE)
#'
#' }
add_gauss_fit_info_to_xslx <- function(excel_file,
                                     excel_sheet_name,
                                     excel_nrows = 100,
                                     excel_ncols = 180,
                                     activate_Plots = FALSE,
                                     select_last_10000 = TRUE,
                                     alpha_packet_size = 100,
                                     activate_save_back_to_xlsx=TRUE){

  info_results <- extract_info_from_xlsx(excel_file)

  gem_type <- info_results$gem_type
  file_locations <- info_results$file_locations

  #Get Data From Excel File
  DataFromSource <- extract_dataFrame_from_xlsx(excel_file = excel_file,
                                        excel_sheet_name = excel_sheet_name,
                                        excel_nrows = excel_nrows)

  for (j in seq(from=1, to=excel_nrows-1)){
    if(DataFromSource$Fit.Gauss.Arrival.Time[j]==0 ||is.na(DataFromSource$Fit.Gauss.Arrival.Time[j])==TRUE){
      cat("\n--- Gauss Fit Calculation for i=",j," skipped.-------------------------------------- \n")
    }else if(DataFromSource$Fit.Gauss.Arrival.Time[j]==1){
      cat("\n--- Gauss Fit Calculation for i=",j," starting.------------------------------------- \n")
      fit_stepsize <- DataFromSource$Fit.Stepsize[j]

      if(TRUE==is.na(DataFromSource$Fit.Stepsize[j])){
        #Default stepsize:
        fit_stepsize = 1000
        DataFromSource$Fit.Stepsize[j] <- fit_stepsize
      }





      if(is.na(DataFromSource$Fit.Nr.of.Curves[j])){
        #Default No of curves:
        fit_nr_curves = 10
        #Check if enough ions were counted:
        if(DataFromSource$No.of.1.Ionizations.per.Trigger[j]<fit_stepsize*fit_nr_curves){
          fit_nr_curves = floor(DataFromSource$No.of.1.Ionizations.per.Trigger[j]/fit_stepsize)
          DataFromSource$Fit.Nr.of.Curves[j] <- fit_nr_curves
        }
      }
      fit_nr_curves <- DataFromSource$Fit.Nr.of.Curves[j]

      if(fit_nr_curves==10){
        select_last_10000=TRUE
      } else{ select_last_10000= FALSE}

      data_arrival_time <- get_data_arrival_time(i=j,
                                                 excel_file = excel_file,
                                                 excel_sheet_name = excel_sheet_name,
                                                 excel_nrows = excel_nrows,
                                                 excel_ncols = excel_ncols,
                                                 select_last_10000 = select_last_10000,
                                                 alpha_packet_size = alpha_packet_size
      )


      fit_guess = c(NA,NA,NA,NA)
      if(is.na(DataFromSource$fit_guess.available[j])){
        DataFromSource$fit_guess.available[j] <- 0}

      if(DataFromSource$fit_guess.available[j]==1){
        fit_guess = c(DataFromSource$fit_guess.Sigma[j],
                      DataFromSource$fit_guess.Mean[j],
                      DataFromSource$`fit_guess.A.(scaling)`[j],
                      DataFromSource$`fit_guess.B.(background)`[j])
      }



      pressure = gsub("\\.", "p",DataFromSource$`Pressure.[torr]`[j])
      plot_name = paste("ArrivalTime_Fit_",gem_type,"_",
                        DataFromSource$`HV.[-V]`[j],"V_",
                        DataFromSource$`Drift.[V]`[j],"V_",
                        pressure,"torr_",
                        excel_sheet_name,"_",
                        "Meas",DataFromSource$Meas.No[j], sep="")
      cat("\n Stepsize: ",fit_stepsize)
      fit_data <- create_gaussian_fits(arrival_time_data=data_arrival_time,
                                  time_window = DataFromSource$`Time.Window.[ms]`[j], #ms
                                  stepsize = fit_stepsize,
                                  meas_name = paste(gem_type,"-",excel_sheet_name,": Meas No. ", DataFromSource$Meas.No[j],sep=""),
                                  activate_Plots = activate_Plots,
                                  plot_name = plot_name,
                                  fit_guess = fit_guess #("Sigma", "Mu", "A", "B")
      )
      DataFromSource$`Fit:.Mean.Arrival.Time`[j] <- mean(fit_data[,2])
      DataFromSource$`Fit:.Sdev(Mean.Arrival.Time)`[j] <- stats::sd(fit_data[,2])
      DataFromSource$`Fit:.Sigma.Arrival.Time`[j] <- mean(fit_data[,1])
      DataFromSource$`Fit:.Sdev(Sigma.Arrival.Time)`[j] <- stats::sd(fit_data[,1])
      cat("\n Gauss fit for i=",j," calculated successfully. \n")
    }

  }
  if(activate_save_back_to_xlsx == TRUE){
    save_dataFrame_back_to_xlsx(excel_file = excel_file,
                                excel_sheet_name = excel_sheet_name,
                                excel_data_frame = DataFromSource)
  }



  cat("\n **************************************************************\n")


  return(DataFromSource)
}
