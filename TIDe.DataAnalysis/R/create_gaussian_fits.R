#' create_gaussian_fits
#'
#' @description This function creates fits based on the function:
#' \deqn{f(x) =B + A * exp(-0.5 * [(x-mu) / sigma)^2] / (sigma * sqrt(2))}
#' It takes a vector filled with arrival times (\code{arrival_time_data}) and divides this vector into subpackets of
#' a fixed number of ions, given by \code{stepsize}. For each subpacket a gaussian fit is run, and the fit
#' parameters are saved into a dataframe, which will be given as an output. In addition to this, different
#' plots will be generated
#' @details If the length of \code{arrival_time_data} is not a multiple of the step length \code{stepsize}, the number of steps is
#' determined by rounding down, resulting in not using the left over part of the end of data. A Warning Message is shown.
#'
#' The Gaussian fit is created by using stas::nls, a non linear least square estimate.
#' @param arrival_time_data Vector with arrival time data, for example using \code{\link{get_data_arrival_time}}. Should be multiple of \code{stepsize}.
#' @param time_window Size of time window, given in ms.
#' @param stepsize Number of ions used per fit. Default: 1000
#' @param meas_name Measurement name, used in plot titles
#' @param activate_Plots if TRUE, plots are created, if FALSE, they are skipped.
#' @param plot_name Plot name used for saving plot files in png format.
#' @param fit_guess Guessed fit parameters c(Sigma, Mu, A, B). If none are given,
#'   default values based on arithmetic mean and standard deviation are used:
#'   c(Sigma, Mu, A, B) = c(sd(data), mean(data),1,0)
#'
#' @return data.frame with the fit parameters (Sigma, Mu, A, B)
#' @import graphics
#' @import viridis
#' @importFrom("stats", "lm", "sd", "nls")
#' @importFrom("grDevices", "dev.off", "png")
#' @export
#'
#' @examples  \dontrun{
#' data <- get_data_arrival_time(i=1,
#'                              excel_file = "em1p5mm_Data_Analysis.xlsx",
#'                              excel_sheet_name = "M1",
#'                              select_last_10000 = TRUE
#'                              )
#' result <- create_gaussian_fits(arrival_time_data,
#'                 time_window = 1,
#'                 stepsize = 1000,
#'                 meas_name = "em1mm_M1_Meas8",
#'                 activate_Plots = TRUE,
#'                 plot_name = "em1mm_1200HV_10V_1torr_M1_Meas8",
#'                 fit_guess = c(NA, NA, NA, NA) #No guess given
#'                 )
#' }
create_gaussian_fits <- function(arrival_time_data,
                            time_window = 1, #ms
                            stepsize = 1000,
                            meas_name = "",
                            activate_Plots = TRUE,
                            plot_name = "test",
                            fit_guess = c(NA, NA, NA, NA) #("Sigma", "Mu", "A", "B")
){

  cat("\n create_gaussian_fits.R started")
  #initialize
  time_window <- time_window
  fit_guess <- fit_guess
  data <- arrival_time_data
  stepsize <- stepsize


  #Check if there are special starting parameters for the gaussian fits given or none (NA).
  #If none are given, arithmetic standard deviation and mean is taken.
  if(is.na(fit_guess[1])){fit_guess[1] <- stats::sd(data)}
  if(is.na(fit_guess[2])){fit_guess[2] <- mean(data)}
  if(is.na(fit_guess[3])){fit_guess[3] <- 1}
  if(is.na(fit_guess[4])){fit_guess[4] <- 0}
  cat("\n Initial Fit guess: ",fit_guess, "\n")

  #Create plot names
  plot_file_name_Mean_Arrival_Time <- paste(plot_name, "_MeanArrivalTime.png",sep="")
  plot_file_name_all_fits <- paste(plot_name, "_allFits.png",sep="")


  data_all <- data

  #Create empty result matrix
  result_matrix <- matrix(data=c(rep(0, length(data_all)/stepsize)),
                          nrow=length(data_all)/stepsize,
                          ncol=4 #sigma, mu, A,B
                          )
  result <- data.frame(result_matrix)
  names(result) <- c("Sigma", "Mu", "A", "B")

  cat("\nLength of data: ",length(data_all), "\n")
  if(length(data_all) %% stepsize!=0){
    cat("WARNING: Length of data is not a multiple of stepsize!")
  }

  for (i in seq(from=1, to=floor(length(data_all)/stepsize))){
    #Select stepsize large block out of data_all
    data <- data_all[((i-1)*stepsize+1):((i)*stepsize)]
    #Evaluate data--------------------------------------------------
    h <- graphics::hist( data,
               xlab = "Arrival time [microseconds]",
               breaks = sqrt(length(data)),
               main = paste("Arrival time histogram from",(i-1)*stepsize+1," to ",(i*stepsize)),
               cex.main = 0.5,
               freq=FALSE)
    df = data.frame(x = h$mids, y = h$density)

    #fit data---------------------------------------------------------
    func_gauss <- function(x, sigma, mu, A, B){
      return(B+A*exp(-0.5*((x-mu)/sigma)^2)/(sigma*sqrt(2)))
    }
    cat("-----------------------------------\nStart of fit calculation for i = ",i,"\n")
    fit_gauss <- stats::nls(y~func_gauss(x, sigma, mu, A,B),data=df,
                     start=list(sigma = fit_guess[1],
                                mu = fit_guess[2],
                                A = fit_guess[3],
                                B=fit_guess[4]),
                     trace=FALSE
    )
    coef= coef(fit_gauss)


    cat(coef)

    #plot data and fit ---------------------------------------------------

    if(activate_Plots==TRUE){
      graphics::plot(df$x, df$y,
           xlim=c(0,time_window*1000),
           xlab = expression(paste("Arrival Time [" ,mu,"s]")),
           ylab = "Frequency",
           main = "Gaussian fit for Arrival Time")
      graphics::curve(func_gauss(x, sigma=coef[1], mu=coef[2], A=coef[3], B=coef[4]), add=TRUE, col="red", lw=2)
      graphics::abline(v=mean(data), col="green")
      graphics::abline(v=mean(data)+stats::sd(data), col="green", lty=2)
      graphics::abline(v=mean(data)-stats::sd(data), col="green", lty=2)
      graphics::abline(v=coef[2], col="blue")
      graphics::abline(v=coef[2]+coef[1], col="blue", lty=2)
      graphics::abline(v=coef[2]-coef[1], col="blue", lty=2)
      graphics::legend("topright", lty=c(1,1,1), col= c("green", "blue"),
             legend = c("Arithmetic", "NLS fit"))
    }
    #Write fit results into "result"-------------------------------
    result[i,] <- coef
  }
  cat("----------------------------------------------\nFits calculated succesfully.\n")

  h <- graphics::hist( data_all,
             xlab = expression(paste("Arrival time [",mu,"s]",sep="")),
             breaks = sqrt(length(data_all)),
             main = paste("Arrival Time  - All Data"),
             cex.main = 0.5
  )
  df_all = data.frame(x = h$mids, y = h$density)

  colorPalette = viridis::viridis(floor(length(data_all)/stepsize))
  color = rep(colorPalette[1],length(seq(from=1, to=floor(length(data_all)/stepsize),by=1)))

  #Create plot showing all fits and data-------------------
  if(activate_Plots==TRUE){
    grDevices::png(filename = plot_file_name_all_fits, width= 1618, height=1000, pointsize=15)
    graphics::par(mar = c(5, 5, 5, 5))
    graphics::plot(df_all$x, df_all$y,
         xlim=c(0,time_window*1000),
         xlab = expression(paste("Arrival Time [" ,mu,"s]")),
         ylab = "Frequency",
         main = paste(meas_name, " Gaussian fit",sep=""),
         pch=1,
         col="black",
         cex= 2.5,
         cex.main = 2.5,
         cex.axis=2,
         cex.lab=2,
         graphics::grid(col = "gray", lty = "dotted"))
    for (k in seq(from=1, to=length(data_all)/stepsize)){
      graphics::curve(func_gauss(x, sigma=result[k,1], mu=result[k,2], A=result[k,3], B=result[k,4]),
            add=TRUE, col=colorPalette[k], lw=2)
    }
    graphics::abline(v= mean(result[,2]),col= "red", lw=1)
    graphics::abline(v= mean(result[,2])+ stats::sd(result[,2]),col= "red", lw=1, lty=2)
    graphics::abline(v= mean(result[,2])- stats::sd(result[,2]),col= "red", lw=1, lty=2)
    if(length(data_all)/stepsize <=10){
      vec_labels = c("i = 1", "i = 2","i = 3","i = 4","i = 5","i = 6","i = 7","i = 8","i = 9","i = 10")

      graphics::legend("topright", lty=c(rep(1,floor(length(data_all/stepsize)))),
             col= colorPalette,
             cex=2,
             lwd=3,
             legend = vec_labels[1:length(colorPalette)])
    }else{
      no_of_colors = length(colorPalette)
      graphics::legend("topright", lty=c(rep(1,floor(length(data_all/stepsize)))),
             col= c(colorPalette[1],
                    colorPalette[floor(no_of_colors/4)],
                    colorPalette[floor(2*no_of_colors/4)],
                    colorPalette[floor(3*no_of_colors/4)],
                    colorPalette[no_of_colors]),
             cex=2,
             lwd=3,
             legend = c("i=1",
                        paste("i=",floor(no_of_colors/4)),
                        paste("i=",floor(2*no_of_colors/4)),
                        paste("i=",floor(3*no_of_colors/4)),
                        paste("i=",no_of_colors))
             )
    }
    grDevices::dev.off()
  }
  #Create plot showing Mean and Sigma of each section plus linear fit-------------------
  x <- seq(from=1, to=floor(length(data_all)/stepsize),by=1)
  avg <-result[,2]
  sdev <- result[,1]
  # linfit <- coef(lm(avg~x, weights = sdev))
  linfit <- coef(stats::lm(avg~x))
  linear_fit <- function(x){
    return(linfit[1]+linfit[2]*x)}
  ylim <- c(min(avg, na.rm = TRUE)-max(sdev, na.rm = TRUE),
            max(avg, na.rm = TRUE)+max(sdev, na.rm = TRUE))

  if(activate_Plots==TRUE){
    grDevices::png(filename = plot_file_name_Mean_Arrival_Time, width= 1618, height=1000, pointsize=15)
    graphics::par(mar = c(5, 5, 5, 5))
    graphics::plot(x, avg,
         ylim = ylim,
         xlab = paste("Measurements - Stepsize " ,stepsize, sep=""),
         ylab = expression(paste("Arrival Time [" ,mu,"s]")),
         main = paste(meas_name," Evolution of Mean Arrival Time ",  sep=""),
         pch = 15,
         col = color,
         cex= 2.5,
         cex.main = 2.5,
         cex.axis=2,
         cex.lab=2,
         graphics::grid(col = "gray", lty = "dotted")
    )
    graphics::abline(h= mean(result[,2]),col= "red", lw=2)
    graphics::abline(h= mean(result[,2])+ stats::sd(result[,2]),col= "red", lw=2, lty=2)
    graphics::abline(h= mean(result[,2])- stats::sd(result[,2]),col= "red", lw=2, lty=2)
    graphics::arrows(x, avg-sdev, x, avg+sdev, length=0.1, angle=90, code=3, col=color, lwd=2)
    graphics::curve(linear_fit(x), col="black", lw=2,lty=1, add=TRUE)
    graphics::legend("bottomright",
           legend=c(paste(round(linfit[1],4),"+ ", round(linfit[2],4),"*x")),
           lty=c(1),
           lwd=c(3),
           cex=1.5,
           col=c("black"))
    grDevices::dev.off()
  }
  return(result)
}
