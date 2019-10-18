## Analisis de Gestion Presupuestal
# Paquetes
library(rio)
library(plotly)
library(dplyr)
library(xlsx)
library(fastmatch)

TRM_REF <- 2850

###
Centros_de_costo_Planta <- import("D:\\Gestion Presupuestal\\Input\\Centros de Costo.xlsx")
Base_Ejecutado <- import("D:\\Gestion Presupuestal\\Input\\Ejecutado\\Presupuesto Ejecutado.xlsx")

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
                               type = 'scatter', mode = 'lines+markers')
for (i in 2:length(LABELS)){
  TSeries_Proveedores <- add_trace(TSeries_Proveedores, x = ~Time_line,
                                   y = Matriz_Costos[,LABELS[i]], name = LABELS[i],
                                   type = 'scatter', mode = 'lines+markers')
}

### Series de tiempo por Servicios por proveedores

Serv_Graphs <- as.list(rep(NA,length(Servicios)))

for (i in 1:length(Serv_Graphs)){
  if (!is.null(dim(Servicios[[i]]))){
    Grafico <- names(Servicios)[i]
    Base <- Servicios[[i]]
    LABELS <- names(Base); LABELS <- LABELS[1:length(LABELS)-1]
    
    TSeries_Proveedores_servicios <- plot_ly(Base, x = ~Time_line,
                                             y = Base[,LABELS[1]], name = LABELS[1], 
                                             type = 'scatter', mode = 'lines+markers')
    if (length(LABELS) > 1){
      for (j in 2:length(LABELS)){
        TSeries_Proveedores_servicios <- add_trace(TSeries_Proveedores_servicios, x = ~Time_line,
                                                   y = Base[,LABELS[j]], name = LABELS[j],
                                                   type = 'scatter', mode = 'lines+markers')
      }
    }
    TSeries_Proveedores_servicios <- layout(TSeries_Proveedores_servicios, title = Grafico)
    Serv_Graphs[[i]] <- TSeries_Proveedores_servicios
  }
}

interes <- sapply(lapply(Serv_Graphs, is.na), all)

Serv_Graphs <- Serv_Graphs[which(!interes)]

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
                         y = Matriz_Costos_Areas[,LABELS[1]], name = LABELS[1],type = 'bar')
                         
for (i in 2:length(LABELS)){
  TSeries_Areas <- add_trace(TSeries_Areas, x = ~Time_line,
                             y = Matriz_Costos_Areas[,LABELS[i]], name = LABELS[i],type = 'bar')
}
layout(TSeries_Areas, yaxis = list(title = 'Count'), barmode = 'stack')


## Distribuible
Distribuible <- Base_Ejecutado[Base_Ejecutado$CC == "Distribuible",]

Proveedor_Distribuible <- unique(Distribuible$Proveedor)

Matriz_Costos_Prov_Dist <- matrix(NA, nrow = length(Time_line), ncol = length(Proveedor_Distribuible))
Matriz_Costos_Prov_Dist <- as.data.frame(Matriz_Costos_Prov_Dist); names(Matriz_Costos_Prov_Dist) <- Proveedor_Distribuible
Matriz_Costos_Prov_Dist$Time_line <- Time_line

for (i in 1:length(Proveedor_Distribuible)){
  Prov_dist <- Proveedor_Distribuible[i]
  Base_por_prov_dist <- Distribuible[Distribuible$Proveedor == Prov_dist,]
  conglomerado <- rowsum(Base_por_prov_dist$Valor, group = Base_por_prov_dist$Time_line)
  Matriz_Costos_Prov_Dist[,i][fmatch(rownames(conglomerado),Matriz_Costos_Prov_Dist$Time_line)] <- conglomerado[,1]
}


### Serie de tiempo por area
LABELS <- names(Matriz_Costos_Prov_Dist); LABELS <- LABELS[1:length(LABELS)-1]

TSeries_Prov_Dist <- plot_ly(Matriz_Costos_Prov_Dist, x = ~Time_line,
                             y = Matriz_Costos_Prov_Dist[,LABELS[1]], name = LABELS[1],type = 'bar')
for (i in 2:length(LABELS)){
  TSeries_Prov_Dist <- add_trace(TSeries_Prov_Dist, x = ~Time_line,
                                 y = Matriz_Costos_Prov_Dist[,LABELS[i]], name = LABELS[i],type = 'bar')
}
layout(TSeries_Prov_Dist, yaxis = list(title = 'Count'), barmode = 'stack')


## Distribuible
Otro <- Base_Ejecutado[Base_Ejecutado$CC == "NA",]

Proveedor_Otro <- unique(Otro$Proveedor)

Matriz_Costos_Prov_Otro <- matrix(NA, nrow = length(Time_line), ncol = length(Proveedor_Otro))
Matriz_Costos_Prov_Otro <- as.data.frame(Matriz_Costos_Prov_Otro); names(Matriz_Costos_Prov_Otro) <- Proveedor_Otro
Matriz_Costos_Prov_Otro$Time_line <- Time_line

for (i in 1:length(Proveedor_Otro)){
  Prov_Otro <- Proveedor_Otro[i]
  Base_por_prov_Otro <- Otro[Otro$Proveedor == Prov_Otro,]
  conglomerado <- rowsum(Base_por_prov_Otro$Valor, group = Base_por_prov_Otro$Time_line)
  Matriz_Costos_Prov_Otro[,i][fmatch(rownames(conglomerado),Matriz_Costos_Prov_Otro$Time_line)] <- conglomerado[,1]
}


### Serie de tiempo por area
LABELS <- names(Matriz_Costos_Prov_Otro); LABELS <- LABELS[1:length(LABELS)-1]

TSeries_Prov_Otro <- plot_ly(Matriz_Costos_Prov_Otro, x = ~Time_line,
                             y = Matriz_Costos_Prov_Otro[,LABELS[1]], name = LABELS[1],type = 'bar')
for (i in 2:length(LABELS)){
  TSeries_Prov_Otro <- add_trace(TSeries_Prov_Otro, x = ~Time_line,
                                 y = Matriz_Costos_Prov_Otro[,LABELS[i]], name = LABELS[i],type = 'bar')
}
layout(TSeries_Prov_Otro, yaxis = list(title = 'Count'), barmode = 'stack')
