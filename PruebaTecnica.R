# Cargar las librerías necesarias para la manipulación de datos, visualización,
# modelos estadísticos complejos, manejo de fechas y matrices.
library(ggplot2)
library(dplyr)
library(lme4)
library(lubridate)
library(Matrix)
library(DBI)
library(odbc)
library(tidyr)

# Función para importar una hoja específica de un archivo Excel como un dataframe
read_excel_sheet <- function(file_path, sheet_name) {
  # Intenta convertir la ruta a una URL para evitar problemas con espacios y caracteres especiales
  excel_file_url <- normalizePath(file_path, winslash = "/", mustWork = TRUE)
  
  # Construye la cadena de conexión ODBC para el archivo Excel
  conn_string <- paste0("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=", excel_file_url, ";")
  
  # Establece y abre la conexión al archivo Excel
  conn <- dbConnect(odbc::odbc(), .connection_string = conn_string)
  
  # Formatea el nombre de la hoja de acuerdo con las convenciones de SQL
  formatted_sheet_name <- paste0("[", sheet_name, "$]")
  
  # Lee los datos de la hoja específica en un dataframe
  df <- dbGetQuery(conn, paste0("SELECT * FROM ", formatted_sheet_name))
  
  # Cierra la conexión al archivo Excel
  dbDisconnect(conn)
  
  return(df)
}

# Rutas a los archivos Excel
municipios_path <- "C:/Users/ccristbo/Dropbox (University of Michigan)/carlos/Work/ADRES/ADRES/Municipios.xlsx"
prestadores_path <- "C:/Users/ccristbo/Dropbox (University of Michigan)/carlos/Work/ADRES/ADRES/Prestadores.xlsx"

# Nombre de la hoja para cada archivo (cambia esto según sea necesario)
municipios_sheet <- "Hoja1"
prestadores_sheet <- "Sheet1"

# Usa la función para leer los datos de las hojas en dataframes
Municipios <- read_excel_sheet(municipios_path, municipios_sheet)
Prestadores <- read_excel_sheet(prestadores_path, prestadores_sheet)

# Renombrar las columnas del data frame 'Prestadores' usando la primera fila de datos.
names(Prestadores) = Prestadores[1,]

# Eliminar la primera fila después de establecerla como nombres de columnas.
Prestadores = Prestadores[-1,]

summary(Prestadores)
summary(Municipios)

# Crear un nuevo data frame 'model' combinando 'Prestadores' con 'Municipios' utilizando
# la clave común 'Depmun', que se extrae del 'codigo_habilitacion'.
model = Prestadores %>%
  mutate(Depmun = substr(codigo_habilitacion, 1, 5)) %>% 
  left_join(Municipios) %>% 
  mutate(clase_personaBi = ifelse(clase_persona=="NATURAL",1,0)) %>% 
  filter(!is.na(Region) & dv>-1) %>% 
  mutate(Region2 = case_when(
    Region =="Región Centro Sur"~1,
    Region =="Región Eje Cafetero"~2,
    Region =="Región Llano"~3,
    Region =="Región Caribe"~4,
    Region =="Región Centro Oriente"~5,
    Region =="Región Pacífico"~6),
    fecha_radicacion = ymd(fecha_radicacion),  # Convertir al formato de fecha
    fecha_vencimiento = ymd(fecha_vencimiento),  # Convertir al formato de fecha
    periodo_validez = as.numeric(difftime(fecha_vencimiento, fecha_radicacion, units = "days")))  

model$dv= as.numeric(model$dv)
# Ajustar un modelo lineal mixto generalizado usando 'glmer' y mostrar el resumen del modelo.
Obj3.fit.1 <- glmer(clase_personaBi ~  periodo_validez+dv+
                      (1 | Region2), family = binomial, data = model)
summary(Obj3.fit.1)
# Extraer y mostrar los efectos aleatorios del modelo ajustado.
random_effects <- ranef(Obj3.fit.1)$Region
random_effects
a = glm(clase_personaBi ~ periodo_validez + dv+Region, family = binomial, data = model)
summary(a)
# Crear dos gráficos usando ggplot2. El primero muestra un gráfico de violín y puntos
# con jitter para 'clase_persona' y 'periodo_validez'. El segundo mejora estéticamente
# el primero y añade colores a los puntos y violines según 'clase_persona', además de
# organizar los paneles de facetas por 'Region'.
plot1 = ggplot(model, aes(x = factor(clase_personaBi), y = periodo_validez, fill = factor(clase_persona))) + 
  geom_violin(trim = FALSE) + 
  scale_fill_manual(values = c("blue", "orange")) + 
  facet_wrap(~ Region, scales = "free_y", ncol = 2) + 
  labs(x = "Clase de Persona", y = "Periodo de Validez") +
  scale_x_discrete(labels = c("0" = "JURÍDICO", "1" = "NATURAL")) + 
  theme_minimal() +
  theme(legend.position = "none", strip.text.x = element_text(angle = 0))


# Guarda el gráfico creado en la ubicación especificada como un archivo JPEG
ggsave("C:/Users/ccristbo/Dropbox (University of Michigan)/carlos/Work/ADRES/ADRES/descripcion1.jpeg", 
       plot1, width = 10, height = 6, dpi = 300)

plot2 = ggplot(model, aes(x = factor(clase_personaBi), y = dv, fill = factor(clase_persona))) + 
  geom_violin(trim = FALSE) + 
  scale_fill_manual(values = c("blue", "orange")) + 
  facet_wrap(~ Region, scales = "free_y", ncol = 2) + 
  labs(x = "Clase de Persona", y = "Periodo de Validez") +
  scale_x_discrete(labels = c("0" = "JURÍDICO", "1" = "NATURAL")) + 
  theme_minimal() +
  theme(legend.position = "none", strip.text.x = element_text(angle = 0))

# Guarda el gráfico creado en la ubicación especificada como un archivo JPEG
ggsave("C:/Users/ccristbo/Dropbox (University of Michigan)/carlos/Work/ADRES/ADRES/descripcion2.jpeg", 
       plot2, width = 10, height = 6, dpi = 300)
