
# Libreria ----------------------------------------------------------------
library(taskscheduleR)



# Tarea -------------------------------------------------------------------
script.refresh <- "C:/Users/CRIST/Desktop/Refresh_Excel/Script_Refresh.R"

taskscheduler_delete("XlRefresh")

taskscheduler_create(taskname = "XlRefresh",rscript = script.refresh,
                     starttime = "20:00",startdate = format(Sys.Date()+1, "%d/%m/%Y"),
                     schedule = "DAILY")

taskcheduler_runnow("XlRefresh")

