
# Libreria ----------------------------------------------------------------
library(RDCOMClient)


# Dashboard ---------------------------------------------------------------
dash <- "C:/Users/CRIST/Desktop/Refresh_Excel/DashBoard.xlsx"
  
## Aplicación de Excel
xl.App <- COMCreate("Excel.Application")

## File
xl.Wbk <- xl.App$Workbooks()$Open(dash)

# Modo Visible
xl.App[["Visible"]] <- TRUE

# Actualización
Sys.sleep(5)
xl.Wbk$RefreshAll()

Sys.sleep(5)

# Guardado
xl.Wbk$Close(FALSE) #TRUE -> para guardar el archivo
xl.App$Quit()
  


