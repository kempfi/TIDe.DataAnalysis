#' make_plot_CpT
#'
#' @description Creates a plot showing the time evolution of the mean counts per trigger.
#'
#' @param meancpa Sum of the rows from \code{mfpga_bin} and divided by alpha packet size.
#' @param meas_no Measurement number.
#' @param alpha_packet_size Number of alpha particles grouped into one data point, default 100.
#'
#' @return renderPlot object
#'
#' @export
#'
#' @importFrom shiny renderPlot
#' @importFrom graphics plot
#' @importFrom graphics abline
#'
#' @examples \dontrun{
#' output$roi_plot_CpT <- make_plot_CpT(meancpa = roi_values$meas$meancpa,
#'                                meas_no = input$roi_meas_no,
#'                                alpha_packet_size = input$roi_alpha_packet_size)
#' }
make_plot_CpT <- function(meancpa,
                          meas_no,
                          alpha_packet_size


){
  meas_no = meas_no
  alpha_packet_size = alpha_packet_size


  CpT_plot <- shiny::renderPlot({
    graphics::plot(1:length(meancpa), meancpa,
         xlab = paste("Consecutive Means (Alpha Packet Size = ",
                      alpha_packet_size, ")", sep = ""),
         ylab = "Mean Counts per Alpha",
         main = paste("Meas",meas_no , ": Mean Counts per Trigger"),
         cex.main = 1
    )
    graphics::abline(h = mean(meancpa), col = "red")
  })



  return(CpT_plot)
}
