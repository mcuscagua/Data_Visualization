
library(rio)
library(fastmatch)
library(plotly)

TRM <- import('TRM.xlsx')
TRM_REF <- 3000

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
  return(paste(as.character(as.numeric(y[2])+1),mes,sep = "-"))
}

trimestre <- function(x){
  Y <- strsplit(x,"-")[[1]]
  if (as.numeric(Y[2]) < 4){
    return(paste(Y[1], "Q1", sep = "-"))
  }else if (as.numeric(Y[2]) < 7){
    return(paste(Y[1], "Q2", sep = "-"))
  }else if (as.numeric(Y[2]) < 10){
    return(paste(Y[1], "Q3", sep = "-"))
  }else{
    return(paste(Y[1], "Q4", sep = "-"))
  }
}

arreglar_mes <- function(x){
  if (x < 10) return(paste("0",as.character(x), sep = "")) else return(as.character(x))
}

### Preprocesamiento de la informacion

TRM$Month <- sapply(TRM$Month, arreglar_mes)
TRM$Time_line <- mapply(function(x,y){return(paste(as.character(x), as.character(y), sep = "-"))}, TRM$Year, TRM$Month)

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

Base_Ejecutado$Proveedor <- Proveedores
Base_Ejecutado$Servicios <- Servicios

Proveedores <- unique(Proveedores)
# Time_line <- unique(Base_Ejecutado$`Mes Pago`); Time_line <- sapply(Time_line, Mes_Pago)

Base_Ejecutado$Time_line <- sapply(Base_Ejecutado$`Mes Pago`,Mes_Pago)

Base_Ejecutado$Quarter <- sapply(Base_Ejecutado$Time_line, trimestre)

Time_line <- sort(unique(Base_Ejecutado$Quarter))

Matriz_Costos <- matrix(NA, nrow = length(unique(Base_Ejecutado$Quarter)), ncol = length(Proveedores))
Matriz_Costos <- as.data.frame(Matriz_Costos); names(Matriz_Costos) <- Proveedores
Matriz_Costos$Time_line <- Time_line

Servicios <- as.list(rep(NA,length(Proveedores))); names(Servicios) <- Proveedores

Base_Ejecutado$Valor[Base_Ejecutado$Moneda == "USD"] <- Base_Ejecutado$Valor[Base_Ejecutado$Moneda == "USD"]*TRM$TRM[fmatch(Base_Ejecutado$Time_line[Base_Ejecutado$Moneda == "USD"], TRM$Time_line)]

for (i in 1:length(Proveedores)){
  proveedor <- Proveedores[i]
  Base_por_proveedor <- Base_Ejecutado[Base_Ejecutado$Proveedor == proveedor,]
  conglomerado <- rowsum(Base_por_proveedor$Valor, group = Base_por_proveedor$Quarter)
  Matriz_Costos[,i][fmatch(rownames(conglomerado),Matriz_Costos$Time_line)] <- conglomerado[,1]
  
  servicios <- unique(Base_por_proveedor$Servicios); servicios <- servicios[!is.na(servicios)]
  
  if (length(servicios) != 0){
    Matriz_Costos_servicios <- matrix(NA, nrow = length(Time_line), ncol = length(servicios))
    Matriz_Costos_servicios <- as.data.frame(Matriz_Costos_servicios); names(Matriz_Costos_servicios) <- servicios
    Matriz_Costos_servicios$Time_line <- Time_line
    
    for (j in 1:length(servicios)){
      cri_serv <- Base_por_proveedor$Servicios == servicios[j]; cri_serv[is.na(cri_serv)] <- F
      Base_por_servicio <- Base_por_proveedor[cri_serv,]
      conglomerado <- rowsum(Base_por_servicio$Valor, group = Base_por_servicio$Quarter)
      Matriz_Costos_servicios[,j][fmatch(rownames(conglomerado),Matriz_Costos_servicios$Time_line)] <- conglomerado[,1]
    }
    Servicios[[i]] <- Matriz_Costos_servicios
  }
}

Servicios$BLOOMBERG


write.csv(Matriz_Costos, "Proveedores_Organizados.csv")

### Serie de tiempo costos totales por proveedores

LABELS <- names(Matriz_Costos); LABELS <- LABELS[1:length(LABELS)-1]

Plot_Prove <- plot_ly(Matriz_Costos, x = ~Time_line,y = Matriz_Costos[,LABELS[1]], name = LABELS[1], type = 'bar')

for (i in 2:length(LABELS)){
  Plot_Prove <- add_trace(Plot_Prove, y = Matriz_Costos[,LABELS[i]], name = LABELS[i])
}
  
Plot_Prove <- layout(Plot_Prove, 
                     xaxis = list(title = 'Tiempo'), 
                     yaxis = list(title = 'Costo (COP)'),
                     barmode = 'stack')


Bloomberg <- Servicios$BLOOMBERG
Bloomberg <- Bloomberg[,1:(ncol(Bloomberg)-1)]
Bloomberg[is.na(Bloomberg)] = 0

Bloomberg_V = sort(apply(Bloomberg, 2, sum), decreasing = T)


Total = Matriz_Costos[,1:(ncol(Matriz_Costos)-1)]
Total[is.na(Total)] = 0

Total_V = sort(apply(Total, 2, sum), decreasing = T)
