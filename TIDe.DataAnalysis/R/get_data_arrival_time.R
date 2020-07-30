#' @title get_data_arrival_time
#' @description This function opens a certain sheet of an Excel file (.xlsx) and analyzes its \code{i}-th row
#' and saves the arrival time data of its ions into a vector. The arrival times are given in ms and there is an
#' option for only selecting the last 10'000 Ions for the calculation. Additionally, a special region of interest
#' can be given by using the parameter \code{select_region}.
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
#' @param i Describes which row of the Excel File shall be analysed. Header does not count as row, so i=1 is the first entry.
#' @param excel_file Excel file name
#' @param excel_sheet_name Name of Excel Sheet
#' @param excel_nrows Number of rows to be copied out, after header (1:(excel_nrows+1)). Default: 100.
#' @param excel_ncols Number of columns to be copied out (1:excel_ncols). Default: 180.
#' @param select_region c(Alpha_Start, Alpha_End). Vector of region of interest for extraction of data arrival
#' @param select_last_10000 If \code{TRUE} only the last 10000 Ions are selected for the Output. If \code{FALSE}
#'   all arrival times are exported.
#' @param alpha_packet_size Number of alpha particles grouped into one data point, default 100.
#'
#' @return vector with arrival times in ms.
#' @export
#' @import graphics
#' @importFrom openxlsx loadWorkbook
#' @importFrom openxlsx read.xlsx
#' @examples \dontrun{
#' data_arrival_time <- get_data_arrival_time(i=1,
#'                              excel_file = "em1p5mm_Data_Analysis.xlsx",
#'                              excel_sheet_name = "M1",
#'                              select_last_10000 = FALSE
#'                              )
#' #Arrival time of first ion:
#' cat("Arrival time of first ion: ", data_arrival_time[1])
#' }
get_data_arrival_time <- function(i=1,
                                  excel_file,
                                  excel_sheet_name,
                                  excel_nrows = 100,
                                  excel_ncols = 180,
                                  select_last_10000 = FALSE,
                                  select_region = NA,
                                  alpha_packet_size = 100

){

  info_results <- extract_info_from_xlsx(excel_file)

  gem_type <- info_results$gem_type
  file_locations <- info_results$file_locations

  DataFromSource <- extract_dataFrame_from_xlsx(excel_file,
                                                excel_sheet_name,
                                                excel_nrows,
                                                excel_ncols)


  #Extract measurement data of Measurement i from binary file
  results <- extract_meas_data_from_xlsx(excel_file,
                                         excel_data_frame = DataFromSource,
                                         meas_no = i,
                                         alpha_packet_size =alpha_packet_size)
  meancpa <- results$meancpa
  alpha_non_zero <- results$alpha_non_zero
  counts_pro_alpha <- results$counts_pro_alpha
  mfpga_bin <- results$mfpga_bin
  alph_start <- results$alph_start
  alph_end <- results$alph_end
  alpha_packet_size <- results$alpha_packet_size

  #---Filter for noise
  filter_out_noise <- DataFromSource[i,159]
  if(is.na(filter_out_noise)){
    filter_out_noise <- 0
    DataFromSource[i,159] <- 0
  }
  #cat("Filter out noise: ", filter_out_noise)

  if(filter_out_noise == 1){
    #filter out noise
    filter_noise_above_x <- DataFromSource[i,160]

    mfpga_bin_df <- data.frame(mfpga_bin)

    #filter out rows where more ioniziations than filer_noise_above_x were recorded

    mfpga_bin_df_filtered <- mfpga_bin_df[mfpga_bin_df[,filter_noise_above_x]==0,]

    #convert df back to matrix
    mfpga_bin <- data.matrix(mfpga_bin_df_filtered)

    #Recalculate results based on mfpga_bin:
    counts_pro_alpha <- rowSums(mfpga_bin[alph_start:alph_end,] !=0)
    alpha_non_zero <- sum(rowSums(mfpga_bin[alph_start:alph_end,]) != 0)
    length_counts_pro_alpha = floor(length(counts_pro_alpha)/alpha_packet_size)*alpha_packet_size
    counts_pro_alpha <- counts_pro_alpha[1:length_counts_pro_alpha]
    mcpa <- matrix(counts_pro_alpha, ncol = alpha_packet_size, byrow = TRUE)
    meancpa <- rowSums(mcpa)/alpha_packet_size
    meancpa[-length(meancpa)] #remove last package because it might only be partially filled

  }

  #Hist
  #Conversion of number entry in mfpga_bin into time:
  #1s = 50'000 -> number in mfpga_bin is divided through 50 and
  #multiplied with time window (zeit_fenster) to get time in ms
  #Manually enter time window
  zeit_fenster <- DataFromSource[i,8]
  graphics::hist( mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]
        [ !mfpga_bin[alph_start:alph_end,1:max(counts_pro_alpha)]==0 ]/(50/zeit_fenster),
        xlab = expression(paste("Arrival time [", mu,"s]",sep="")),
        breaks = 25,
        main = "Arrival Time",
        cex.main = 1)
  data_arrival_time <-  mfpga_bin[,1:max(counts_pro_alpha)][ !mfpga_bin[,1:max(counts_pro_alpha)]==0 ]/(50/zeit_fenster)
  if(select_last_10000==TRUE){

    #Select last 10000 ion arrival times
    data_arrival_time <-data_arrival_time[(length(data_arrival_time)-9999):length(data_arrival_time)]

  }
  cat("Arrival time data extracted succesfully.")
  return(data_arrival_time)
}
