# Paquetes
##########

library(shiny)
library(rio)
library(DT)
library(data.table)
library(rhandsontable)
library(shinythemes)
library(shinydashboard)
library(XML)
library(httr)
library(rvest)
library(fastmatch)
library(plotly)
library(dplyr)

# Funciones
###########


ui <- dashboardPage(
  dashboardHeader(title = div(img(src="logo.png", height = 30), "Control de Presupuesto"),
                  titleWidth = 400),
  dashboardSidebar(
    textOutput("explica")
  ),
  dashboardBody(
    tabsetPanel(
      tabPanel("Proveedores", plotlyOutput("plot_proveedor")),
      tabPanel("Cobros a areas", plotlyOutput("plot_exclusivos"))
    )
  ),
  skin = "black"
)


server <- function(input, output){
  TRM_REF <- 2775
  
  ###
  Centros_de_costo_Planta <- import("Centros de Costo.xlsx")
  Base_Ejecutado <- import("Presupuesto Ejecutado.xlsx")
  
  names(Base_Ejecutado) <- c("Linea","ProvServ","Amortizable",
                             "Nro Factura","Posicion","Pedido",
                             "Migo","Moneda","Valor",
                             "Tasa","IVA","Valor Total",
                             "Retencion","Pago Neto","Fecha Llegada",
                             "F_Entrega","F_Cont","F_Pago",
                             "F_Env_Swift","F_Env_Cert_Ret","CC",
                             "Cod Material","Receptor","Mes Pago",
                             "Mes Causa","Observaciones","Amort_1",
                             "Amort_2","Amort_3","Label",
                             "F_Sol_Cert_Ret","Hora Reg","Responsable",
                             "Pos Actual","F Prox Factura","Val Prox Factura",
                             "PPP","Valor Sin IVA","X1",
                             "PPMP","Correo Migo")
  
  ### Funciones auxiliares
  Mes_Pago <- function(x){
    y <- strsplit(x,"-")[[1]]
    mes <- switch(y[1],
                  ENERO = "01",
                  FEBRERO = "02",
                  MARZO = "03",
                  ABRIL = "04",
                  MAYO = "05",
                  JUNIO = "06",
                  JULIO = "07",
                  AGOSTO = "08",
                  SEPTIEMBRE = "09",
                  OCTUBRE = "10",
                  NOVIEMBRE = "11",
                  DICIEMBRE = "12")
    return(paste(y[2],mes,sep = "-"))
  }
  
  ### Preprocesamiento de la informacion
  Base_Ejecutado$CC[is.na(Base_Ejecutado$CC)] <- "NA"
  
  Proveedores <- sapply(Base_Ejecutado$ProvServ,function(x){
    Proveedor <- strsplit(x," - ")[[1]][1]
    if (grepl("BVC", Proveedor)){
      aux <- strsplit(x, " ")[[1]]
      Proveedor <- aux[1]
    }
    if (grepl("SET ICAP FX", Proveedor)){
      aux <- strsplit(x, " ")[[1]]
      Proveedor <- paste(aux[1:3],collapse = " ")
    }
    if (grepl("FIDESSA", Proveedor)){
      aux <- strsplit(x, " ")[[1]]
      Proveedor <- aux[1]
    }
    if (grepl("AMV", Proveedor)){
      aux <- strsplit(x, " ")[[1]]
      Proveedor <- aux[1]
    }
    return(Proveedor)
  })
  Servicios <- sapply(Base_Ejecutado$ProvServ,function(x){
    Servicio <- strsplit(x," - ")[[1]][2]
    if (grepl("BVC", x)){
      aux <- strsplit(x, " ")[[1]]
      Servicio <- paste(aux[2:length(aux)], collapse = " ")
    }
    if (grepl("SET ICAP FX", x)){
      aux <- strsplit(x, " ")[[1]]
      Servicio <- paste(aux[4:length(aux)], collapse = " ")
    }
    if (grepl("FIDESSA", x)){
      aux <- strsplit(x, " ")[[1]]
      Servicio <- paste(aux[2:length(aux)], collapse = " ")
    }
    if (grepl("AMV", x)){
      aux <- strsplit(x, " ")[[1]]
      Servicio <- paste(aux[2:length(aux)], collapse = " ")
    }
    return(Servicio)
  })
  
  Base_Ejecutado$Valor[Base_Ejecutado$Moneda == "USD"] <- Base_Ejecutado$Valor[Base_Ejecutado$Moneda == "USD"]*TRM_REF
  
  Base_Ejecutado$Proveedor <- Proveedores
  Base_Ejecutado$Servicios <- Servicios
  
  Proveedores <- unique(Proveedores)
  Time_line <- unique(Base_Ejecutado$`Mes Pago`); Time_line <- sapply(Time_line, Mes_Pago)
  Time_line <- sort(Time_line)
  Base_Ejecutado$Time_line <- sapply(Base_Ejecutado$`Mes Pago`,Mes_Pago)
  
  Matriz_Costos <- matrix(NA, nrow = length(Time_line), ncol = length(Proveedores))
  Matriz_Costos <- as.data.frame(Matriz_Costos); names(Matriz_Costos) <- Proveedores
  Matriz_Costos$Time_line <- Time_line
  
  Servicios <- as.list(rep(NA,length(Proveedores))); names(Servicios) <- Proveedores
  
  for (i in 1:length(Proveedores)){
    proveedor <- Proveedores[i]
    Base_por_proveedor <- Base_Ejecutado[Base_Ejecutado$Proveedor == proveedor,]
    conglomerado <- rowsum(Base_por_proveedor$Valor, group = Base_por_proveedor$Time_line)
    Matriz_Costos[,i][fmatch(rownames(conglomerado),Matriz_Costos$Time_line)] <- conglomerado[,1]
    
    servicios <- unique(Base_por_proveedor$Servicios); servicios <- servicios[!is.na(servicios)]
    
    if (length(servicios) != 0){
      Matriz_Costos_servicios <- matrix(NA, nrow = length(Time_line), ncol = length(servicios))
      Matriz_Costos_servicios <- as.data.frame(Matriz_Costos_servicios); names(Matriz_Costos_servicios) <- servicios
      Matriz_Costos_servicios$Time_line <- Time_line
      
      for (j in 1:length(servicios)){
        cri_serv <- Base_por_proveedor$Servicios == servicios[j]; cri_serv[is.na(cri_serv)] <- F
        Base_por_servicio <- Base_por_proveedor[cri_serv,]
        conglomerado <- rowsum(Base_por_servicio$Valor, group = Base_por_servicio$Time_line)
        Matriz_Costos_servicios[,j][fmatch(rownames(conglomerado),Matriz_Costos_servicios$Time_line)] <- conglomerado[,1]
      }
      Servicios[[i]] <- Matriz_Costos_servicios
    }
  }
  
  
  ### Serie de tiempo costos totales por proveedores
  
  LABELS <- names(Matriz_Costos); LABELS <- LABELS[1:length(LABELS)-1]
  
  TSeries_Proveedores <- plot_ly(Matriz_Costos, x = ~Time_line,
                                 y = Matriz_Costos[,LABELS[1]], name = LABELS[1], 
                                 type = 'scatter', mode = 'lines+markers') %>%
    layout(xaxis = list(title = NA),
           yaxis = list(title = 'Costo (COP)'))
  
  for (i in 2:length(LABELS)){
    TSeries_Proveedores <- add_trace(TSeries_Proveedores, x = ~Time_line,
                                     y = Matriz_Costos[,LABELS[i]], name = LABELS[i],
                                     type = 'scatter', mode = 'lines+markers')
  }
  
  ### Segregacion por tipos de cobro: Exclusivo, Distribuible, Otro
  
  ## Exclusivo
  Exclusivo <- Base_Ejecutado[!(Base_Ejecutado$CC %in% c("NA", "Distribuible")),]
  Exclusivo$Unidad_Organizativa <- Centros_de_costo_Planta$`UNIDAD ORGANIZATIVA`[fmatch(Exclusivo$CC,Centros_de_costo_Planta$`CENTRO DE COSTOS`)]
  
  Exclusivo_Revisar <- Exclusivo[is.na(Exclusivo$Unidad_Organizativa),]
  Exclusivo <- Exclusivo[!is.na(Exclusivo$Unidad_Organizativa),]
  
  Areas <- unique(Exclusivo$Unidad_Organizativa)
  
  Matriz_Costos_Areas <- matrix(NA, nrow = length(Time_line), ncol = length(Areas))
  Matriz_Costos_Areas <- as.data.frame(Matriz_Costos_Areas); names(Matriz_Costos_Areas) <- Areas
  Matriz_Costos_Areas$Time_line <- Time_line
  
  for (i in 1:length(Areas)){
    Area <- Areas[i]
    Base_por_area <- Exclusivo[Exclusivo$Unidad_Organizativa == Area,]
    conglomerado <- rowsum(Base_por_area$Valor, group = Base_por_area$Time_line)
    Matriz_Costos_Areas[,i][fmatch(rownames(conglomerado),Matriz_Costos_Areas$Time_line)] <- conglomerado[,1]
  }
  
  
  ### Serie de tiempo por area
  LABELS <- names(Matriz_Costos_Areas); LABELS <- LABELS[1:length(LABELS)-1]
  
  TSeries_Areas <- plot_ly(Matriz_Costos_Areas, x = ~Time_line,
                           y = Matriz_Costos_Areas[,LABELS[1]], name = LABELS[1],type = 'bar') %>%
    layout(xaxis = list(title = NA),
           yaxis = list(title = 'Costo (COP)'),
           barmode = 'stack')
  
  for (i in 2:length(LABELS)){
    TSeries_Areas <- add_trace(TSeries_Areas, x = ~Time_line,
                               y = Matriz_Costos_Areas[,LABELS[i]], name = LABELS[i],type = 'bar')
  }
  
  
  ##### Rendering
  
  output$explica <- renderText({
    texto <- "El informe presenta el tablero de control para la administracion del presupuesto para la gestoria transaccional"
    return(texto)
  })
  
  output$plot_proveedor <- renderPlotly({
    return(TSeries_Proveedores)
  })
  
  output$plot_exclusivos <- renderPlotly({
    return(TSeries_Areas)
  })
  
}

shinyApp(ui = ui, server = server)
