#' extract_meas_data_from_xlsx
#'
#' @description This function extracts the measurement data of a single measurement in a
#' dataframe. It uses \code{\link{extract_info_from_xlsx}}.
#'
#' @param excel_file Excel File Name as a string, example: "Data_Analysis.xlsx" If the file is
#' in a different folder use the datapath, example: "C:/Users/username/Documents/R_Coding/Excel_File.xlsx"
#' @param excel_data_frame Data Frame, obtained using \code{\link{extract_dataFrame_from_xlsx}}.
#' @param meas_no Measurement number
#' @param alpha_packet_size Number of alpha particles to be grouped together into one data point,
#' default 100.
#' @param activate_darkRate_Mode Default False. If this is activated the dark rate measurement data is extracted
#' instead of the regular measurement data.
#'
#' @details Excel sheet must contain a sheet named "Info", where in cell B1 the
#'   THGEM type of the measurements is given and in cell B2 the path to the
#'   folder with measurement data is given. In the general measurement data folder,
#'   subfolders for each THGEM should be present with the THGEM name as folder name.
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
#' @return Returns a list list with \code{meancpa}, \code{alpha_non_zero}, \code{counts_pro_alpha},
#' \code{mfpga_bin},\code{alph_start},\code{alph_end} and \code{alpha_packet_size}.
#'
#' @export
#'
#' @examples \dontrun{
#' roi_values$meas <- extract_meas_data_from_xlsx(excel_file = roi_values$roi_xlsx_file_datapath,
#'                              excel_data_frame = roi_values$data,
#'                              meas_no = input$roi_meas_no,
#'                              alpha_packet_size = input$roi_alpha_packet_size)
#'
#' meancpa <- roi_values$meas$meancpa
#' alpha_packet_size <- input$roi_alpha_packet_size
#'
#' }
extract_meas_data_from_xlsx<- function(excel_file,
                            excel_data_frame,
                            meas_no,
                            alpha_packet_size =100,
                            activate_darkRate_Mode = FALSE){

  #Open Excel Sheet Info and extract gem type and file locations
  data_info <- extract_info_from_xlsx(excel_file = excel_file)

  file_locations <- data_info$file_locations
  gem_type <- data_info$gem_type


  DataFromSource <- excel_data_frame
  i <- meas_no
  cat("\n i: ", i)

  if(activate_darkRate_Mode == TRUE){
    if(DataFromSource$DR.File[i] == "NONE"){
      cat("No DR file.")
    }else{
      if(DataFromSource$DR.File[i] == "OTHER"){
        cat("Other DR file.")
        file_name <- DataFromSource$DR.Other.file.name[i]
      }else{
        #Read in filename from Excel sheet:
        file_name <- paste(DataFromSource[i,2],".bin",sep="")
        file_name <-paste(DataFromSource[i,79], file_name, sep="_") #add dr_prefix
      }
  }
  }else{
    file_name <- paste(DataFromSource[i,2],".bin",sep="")
  }



  #Create data_file by adding location to filename
  data_file <- paste(file_locations, "/", file_name, sep = "")

  cat("\n File Name: ", file_name)
  #cat("\n Data File: ", data_file)

  #Read binary file into data_fpga_bin.
  #Remove all start package "65535" numbers so that we are left with ONLY our measurement data.
  datei_groesse <- file.size(data_file)
  data_fpga_bin <- readBin(con=data_file, what="int", n=datei_groesse, 2, signed = FALSE)
  start_package <- which(data_fpga_bin %in% 65535)
  data_fpga_bin <- data_fpga_bin[-start_package]

  #Convert mfpga_bin Into a matrix with 32 columns.
  #Each row corresponds to one alpha, each column to an ionization.
  #Numbers inside correspond to time (1s=50'000)
  mfpga_bin <-  matrix(data_fpga_bin, ncol = 32, byrow = T)
  mfpga_bin <- mfpga_bin[-nrow(mfpga_bin),] #Throw away last row


  #Define starting and ending point by a default value and
  #overwrite if something is written in Excel file
  if(DataFromSource$Alpha.Start[i]=="TBD"|| is.na(DataFromSource$Alpha.Start[i])){
    alph_start <- 1
    #print("Alpha Start TBD, default value = 1.")
  }else{
    alph_start <- as.numeric(DataFromSource$Alpha.Start[i])
  }
  if(DataFromSource$Alpha.End[i]=="TBD"|| is.na(DataFromSource$Alpha.End[i])){
    alph_end <- nrow(mfpga_bin)
    #print("Alpha End TBD, default value = Total No of Alphas.")
  }else{
    alph_end <- as.numeric(DataFromSource$Alpha.End[i])
  }

  #Sum all non zero entries from mfpga_bin to give a vector with how many ionizations per alpha.
  #No Columns = No of rows from mfpga_bin = Total no of alphas
  counts_pro_alpha <- rowSums(mfpga_bin[alph_start:alph_end,] !=0)  #Type: Double
  #cat("Counts pro alpha: ",counts_pro_alpha)


  #Define alpha packet size
  alpha_packet_size = alpha_packet_size

  #Count all rows with at least one ionization, save number as alpha_non_zero
  alpha_non_zero <- sum(rowSums(mfpga_bin[alph_start:alph_end,]) != 0)
  #cat("Alpha Non Zero: ",alpha_non_zero)

  #Cut off end of counts_pro_alpha, so that all packets are full.
  length_counts_pro_alpha = floor(length(counts_pro_alpha)/alpha_packet_size)*alpha_packet_size
  counts_pro_alpha <- counts_pro_alpha[1:length_counts_pro_alpha]


  mcpa <- matrix(counts_pro_alpha, ncol = alpha_packet_size, byrow = TRUE)
  meancpa <- rowSums(mcpa)/alpha_packet_size
  meancpa[-length(meancpa)] #remove last package because it might only be partially filled

  results <- list("meancpa" = meancpa,
                  "alpha_non_zero" = alpha_non_zero,
                  "counts_pro_alpha" = counts_pro_alpha,
                  "mfpga_bin" = mfpga_bin,
                  "alph_start" = alph_start,
                  "alph_end" = alph_end,
                  "alpha_packet_size" = alpha_packet_size)

  return(results)
}







