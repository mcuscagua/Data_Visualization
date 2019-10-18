library(rio)
library(fastmatch)
library(plotly)

Matriz_Costos <- import('Proveedores_Organizados.xlsx')

LABELS <- names(Matriz_Costos); LABELS <- LABELS[2:length(LABELS)]

Plot_Prove <- plot_ly(Matriz_Costos, x = ~Time_line, y = Matriz_Costos[,LABELS[1]], name = LABELS[1], type = 'bar')

for (i in 2:length(LABELS)){
  Plot_Prove <- add_trace(Plot_Prove, y = Matriz_Costos[,LABELS[i]], name = LABELS[i])
}

Plot_Prove <- layout(Plot_Prove, 
                     xaxis = list(title = 'Tiempo'), 
                     yaxis = list(title = 'Costo (COP)'),
                     barmode = 'stack')
plotly_IMAGE(Plot_Prove, format = "png", out_file = "Mes_a_Mes.png", width = 1000, height = 1000)
