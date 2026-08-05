#Paquetes
library(readr)
library(dplyr)
library(tidyr)
library(writexl)
library(readxl)
library(flextable)
library(officer)
library(purrr)
library(openxlsx)
#Depuración
IPS <- read_csv("Datos_Brutos/IPS.csv")
mental_health <- c("Salud Mental Adulto", "Salud Mental Pediátrico", 
                   "Salud Mental", "Psiquiatría", "Cuidado Agudo Mental")
spa <- c("SPA", "SPA Básico Adultos", "SPA Básico Pediátricos", 
         "SPA Adultos", "SPA Pediátricas", "Farmacodependencia")
resultado_final <- IPS %>%
  filter(`nom descripcion capacidad` %in% c(mental_health, spa)) %>%
  mutate(servicio = if_else(`nom descripcion capacidad` %in% mental_health, 
                            "Salud_Mental", "SPA")) %>%
  group_by(Departamento, servicio, `nom grupo capacidad`) %>%
  summarise(Total = sum(`num cantidad capacidad instalada`, na.rm = TRUE), 
            .groups = "drop") %>%
  unite("variable", servicio, `nom grupo capacidad`, sep = "_") %>%

  pivot_wider(names_from = variable, values_from = Total, values_fill = 0)
print(resultado_final)
write_xlsx(resultado_final, "totales_salud_mental_spa_por_departamento.xlsx")
#Base_Datos_depurada
COL_2023_SM <- read_excel("Data_Depurada/BASE_COL_SALMENT_2023.xlsx")
#TABLA 1 Articulo
vars <- c(
  "Psiquiatras"                       = "PSIQUI",
  "Psicólogos"                        = "PSICO",
  "Camas salud mental"                = "Cama_SM",
  "Camillas salud mental"             = "Camillas_SM",
  "Sillas salud mental"               = "Sillas_SM",
  "Camas sustancias psicoactivas"     = "Camas_SPA",
  "Camillas sustancias psicoactivas"  = "Camillas_SPA",
  "Sillas sustancias psicoactivas"    = "Sillas_SPA"
)

tabla1 <- map2_dfr(names(vars), vars, function(nombre, col) {
  
  # Ausencia de reporte = 0 oferta, no dato perdido -> replace_na()
  tasa <- replace_na(COL_2023_SM[[col]], 0) / COL_2023_SM$Pobla_Total * 100000
  dep  <- COL_2023_SM$Departamento
  
  media   <- mean(tasa)
  mediana <- median(tasa)
  
  encima <- dep[tasa > media]
  encima <- encima[order(-tasa[tasa > media])]
  
  sin_recurso_dep <- dep[tasa == 0]
  sin_recurso_txt <- if (length(sin_recurso_dep) == 0) {
    "0"
  } else {
    paste0(length(sin_recurso_dep), " (", paste(sin_recurso_dep, collapse = ", "), ")")
  }
  
  tibble(
    Variable = nombre,
    Media = round(media, 2),
    Mediana = round(mediana, 2),
    `Departamentos por encima de la media nacional` = paste(encima, collapse = ", "),
    `Departamentos sin recurso` = sin_recurso_txt
  )
})

ft <- flextable(tabla1) %>%
  set_table_properties(layout = "autofit") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  fontsize(size = 10, part = "all") %>%
  bold(part = "header") %>%
  align(j = 2:3, align = "center", part = "all") %>%
  border_remove() %>%
  hline_top(part = "header", border = fp_border(width = 1.5, color = "black")) %>%
  hline_bottom(part = "header", border = fp_border(width = 1, color = "black")) %>%
  hline_bottom(part = "body", border = fp_border(width = 1.5, color = "black")) %>%
  padding(padding = 3, part = "all")

nota_txt <- paste0(
  "Nota. Las tasas de psiquiatras, psicólogos, camas, camillas y sillas se expresan por 100.000 ",
  "habitantes. \"Departamentos sin recurso\" corresponde al número de departamentos que no ",
  "reportaron disponibilidad del recurso evaluado."
)

doc <- read_docx() %>%
  body_add_fpar(fpar(ftext("Tabla 1", fp_text(bold = TRUE, font.family = "Times New Roman")))) %>%
  body_add_fpar(fpar(ftext("Descripción de la oferta de salud mental por departamentos, Colombia, 2023",
                           fp_text(italic = TRUE, font.family = "Times New Roman")))) %>%
  body_add_flextable(ft) %>%
  body_add_fpar(fpar(ftext(nota_txt, fp_text(italic = TRUE, font.family = "Times New Roman", font.size = 9))))

print(doc, target = "Tabla1_Oferta_SaludMental_APA.docx")

# =========================================================
# 1. Definición de grupos de variables
# =========================================================

# Demanda - RIPS (11 grupos CIE-10) y SIVIGILA (intento suicida), por sexo
diag_muj <- c("F00-F09_MUJ","F10-F19_MUJ","F20-F29_MUJ","F30-F39_MUJ","F40-F48_MUJ",
              "F50-F59_MUJ","F60-F69_MUJ","F70-F79_MUJ","F80-F89_MUJ","F90-F98_MUJ","F99_MUJ")
diag_hom <- c("F00-F09_HOM","F10-F19_HOM","F20-F29_HOM","F30-F39 _HOM","F40-F48_HOM",
              "F50-F59_HOM","F60-F69_HOM","F70-F79_HOM","F80-F89_HOM","F90-F98_HOM","F99_HOM")
sivigila_muj <- c("INT_SUC_PER__MUJ","INT_SUC_CAS__MUJ")
sivigila_hom <- c("INT_SUC_PER_HOM","INT_SUC_CAS_HOM")

# Oferta - RETHUS y REPS, con población total
oferta_rethus <- c("PSIQUI","PSICO")
oferta_reps   <- c("Camas_SPA","Sillas_SPA","Camillas_SPA","Cama_SM","Sillas_SM","Camillas_SM")

# =========================================================
# 2. Función auxiliar: nombre limpio de la variable de tasa
# =========================================================
clean_name <- function(x) {
  x <- gsub(" _HOM", "_HOM", x)   # corrige el espacio suelto de "F30-F39 _HOM"
  x <- gsub("__MUJ", "_MUJ", x)   # corrige doble guion bajo de INT_SUC_*
  x
}

# =========================================================
# 3. Construcción de la base de tasas
# =========================================================
base_tasas <- COL_2023_SM %>%
  transmute(
    `Código DANE` = `Código DANE`,
    Departamento  = Departamento,
    
    # Demanda mujeres: tasa x 100.000 mujeres
    across(all_of(diag_muj), ~ .x / Pobla_MUJ * 1e5,
           .names = "TASA_{clean_name(.col)}"),
    across(all_of(sivigila_muj), ~ .x / Pobla_MUJ * 1e5,
           .names = "TASA_{clean_name(.col)}"),
    
    # Demanda hombres: tasa x 100.000 hombres
    across(all_of(diag_hom), ~ .x / Pobla_HOM * 1e5,
           .names = "TASA_{clean_name(.col)}"),
    across(all_of(sivigila_hom), ~ .x / Pobla_HOM * 1e5,
           .names = "TASA_{clean_name(.col)}"),
    
    # Oferta: tasa x 100.000 población total.
    # Faltantes en REPS = sin oferta instalada -> se tratan como 0, no como NA
    across(all_of(oferta_rethus), ~ .x / Pobla_Total * 1e5,
           .names = "TASA_{.col}"),
    across(all_of(oferta_reps), ~ tidyr::replace_na(.x, 0) / Pobla_Total * 1e5,
           .names = "TASA_{.col}")
  ) %>%
  arrange(Departamento)
write.xlsx(base_tasas, "BASE_TASAS_SALUD_MENTAL_2023.xlsx",
           sheetName = "BASE_TASAS", overwrite = TRUE)
#BASE_TASAS
TAS_SALMENTAL_2023<- read_excel("Data_Depurada/BASE_TASAS_SALUD_MENTAL_2023.xlsx")
library(tidyverse)
library(lme4)
library(lmerTest)
library(effectsize)
library(performance)
library(emmeans)

# ---------------------------------------------------------------------------
# 1. Reconstruir tabla_tasas (formato largo) desde TAS_SALMENTAL_2023
#    Solo se usan las 11 columnas de tasas por grupo diagnóstico (MUJ/HOM);
#    se excluyen intento de suicidio, psiquiatras, psicólogos y camas.
# ---------------------------------------------------------------------------
tabla_tasas <- TAS_SALMENTAL_2023 %>%
  select(Departamento,
         matches("^TASA_F\\d{2}(-F\\d{2})?_(MUJ|HOM)$")) %>%
  pivot_longer(
    cols = -Departamento,
    names_to      = c("Diagnostico", "Sexo"),
    names_pattern = "TASA_(F\\d{2}(?:-F\\d{2})?)_(MUJ|HOM)",
    values_to     = "Tasa"
  ) %>%
  mutate(
    Sexo = recode(Sexo, MUJ = "Mujeres", HOM = "Hombres"),
    Diagnostico = factor(
      Diagnostico,
      levels = c("F00-F09","F10-F19","F20-F29","F30-F39","F40-F48",
                 "F50-F59","F60-F69","F70-F79","F80-F89","F90-F98","F99"),
      labels = c(
        "F00–F09 Trastornos mentales orgánicos",
        "F10–F19 Substancias psicotrópicas",
        "F20–F29 Esquizofrenia y delirantes",
        "F30–F39 Estado de ánimo",
        "F40–F49 Neuróticos y estrés",
        "F50–F59 Comportamiento",
        "F60–F69 Personalidad",
        "F70–F79 Retraso mental",
        "F80–F89 Desarrollo psicológico",
        "F90–F98 Emocionales en niñez",
        "F99 Sin especificar"
      )
    ),
    Sexo = factor(Sexo, levels = c("Hombres","Mujeres")),
    Departamento = factor(Departamento)
  )
dim(tabla_tasas)
str(tabla_tasas)
options(contrasts = c("contr.sum", "contr.poly"))  # necesario para SS tipo III
modelo_dep <- lmer(
  Tasa ~ Sexo * Diagnostico + (1 | Departamento),
  data = tabla_tasas,
  REML = TRUE
)
anova(modelo_dep, type = 3)
eta_squared(modelo_dep, partial = TRUE)
r2(modelo_dep)
emmeans_sexo_diag <- emmeans(
  modelo_dep,
  ~ Sexo | Diagnostico
)

contrastes_sexo <- contrast(
  emmeans_sexo_diag,
  method = "pairwise",
  adjust = "BY"
)
medias <- as.data.frame(emmeans_sexo_diag) %>%
  select(Diagnostico, Sexo, emmean) %>%
  pivot_wider(names_from = Sexo, values_from = emmean)

pvalores <- as.data.frame(contrastes_sexo) %>%
  select(Diagnostico, p.value)
tabla2 <- left_join(medias, pvalores, by = "Diagnostico") %>%
  rename(p = p.value)
print(tabla2, n = Inf)

# ---------------------------------------------------------------------------
# 4. Formatear valores en estilo APA
#    - decimales consistentes
#    - p < .001 sin cero inicial, según APA 7
#    - asteriscos de significancia para la nota al pie
# ---------------------------------------------------------------------------
formatear_p_apa <- function(p) {
  case_when(
    p < .001 ~ "< .001",
    TRUE     ~ sub("^0", "", sprintf("%.3f", p))  # quita el cero inicial: .038 en vez de 0.038
  )
}

asterisco_sig <- function(p) {
  case_when(
    p < .001 ~ "***",
    p < .01  ~ "**",
    p < .05  ~ "*",
    TRUE     ~ ""
  )
}

tabla2_apa <- tabla2 %>%
  mutate(
    Hombres_fmt = sprintf("%.2f", Hombres),
    Mujeres_fmt = sprintf("%.2f", Mujeres),
    p_fmt       = paste0(formatear_p_apa(p), asterisco_sig(p))
  ) %>%
  select(Diagnostico, Hombres_fmt, Mujeres_fmt, p_fmt) %>%
  rename(
    "Grupo diagnóstico CIE-10" = Diagnostico,
    "Hombres"                  = Hombres_fmt,
    "Mujeres"                  = Mujeres_fmt,
    "p"                        = p_fmt
  )

# ---------------------------------------------------------------------------
# 5. Construir la tabla con formato APA 7 usando flextable
#    - Times New Roman 12pt
#    - Solo líneas horizontales (arriba, bajo encabezado, abajo)
#    - Número de tabla en negrita, título en cursiva, encima de la tabla
#    - Nota explicativa en cursiva, debajo de la tabla
# ---------------------------------------------------------------------------
ft <- flextable(tabla2_apa) %>%
  set_table_properties(layout = "autofit") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  fontsize(size = 12, part = "all") %>%
  bold(part = "header") %>%
  align(align = "center", part = "header") %>%
  align(j = c("Hombres", "Mujeres", "p"), align = "center", part = "body") %>%
  align(j = "Grupo diagnóstico CIE-10", align = "left", part = "body") %>%
  border_remove() %>%
  hline_top(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline_bottom(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline(i = 1, part = "header", border = fp_border(width = 1, color = "black")) %>%
  padding(padding = 4, part = "all")

# ---------------------------------------------------------------------------
# 6. Ensamblar el documento Word: número de tabla + título + tabla + nota
#    (formato APA 7: "Tabla N" en negrita, título en cursiva debajo,
#     nota al pie en cursiva solo en la palabra "Nota.")
# ---------------------------------------------------------------------------
numero_tabla <- fpar(
  ftext("Tabla 2", prop = fp_text(font.family = "Times New Roman", font.size = 12, bold = TRUE))
)

titulo_tabla <- fpar(
  ftext(
    "Medias marginales estimadas por sexo según grupo diagnóstico CIE-10, Colombia, 2023",
    prop = fp_text(font.family = "Times New Roman", font.size = 12, italic = TRUE)
  )
)

nota_tabla <- fpar(
  ftext("Nota. ", prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = TRUE)),
  ftext(
    "Las medias marginales corresponden al modelo mixto (Sexo \u00d7 Diagnóstico, con Departamento como efecto aleatorio). Los valores de p se ajustaron con la corrección de Benjamini y Yekutieli para comparaciones múltiples.",
    prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = FALSE)
  )
)

nota_sig <- fpar(
  ftext(
    "*p < .05. **p < .01. ***p < .001.",
    prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = TRUE)
  )
)

doc <- read_docx() %>%
  body_add_fpar(numero_tabla) %>%
  body_add_fpar(titulo_tabla) %>%
  body_add_flextable(ft) %>%
  body_add_fpar(nota_tabla) %>%
  body_add_fpar(nota_sig)
print(doc, target = "Tabla2_APA.docx")
#tabla3

# ---------------------------------------------------------------------------
# 1. Reconstruir datos de intento de suicidio (formato largo) por sexo
#    NOTA: solo hay tasas (por 100.000 hab.) en TAS_SALMENTAL_2023,
#    no conteos absolutos de "Casos" ni "Personas".
# ---------------------------------------------------------------------------
suicidio_long <- TAS_SALMENTAL_2023 %>%
  select(Departamento,
         TASA_INT_SUC_CAS_MUJ, TASA_INT_SUC_PER_MUJ,
         TASA_INT_SUC_CAS_HOM, TASA_INT_SUC_PER_HOM) %>%
  pivot_longer(
    cols = -Departamento,
    names_to = c("Metrica", "Sexo"),
    names_pattern = "TASA_INT_SUC_(CAS|PER)_(MUJ|HOM)",
    values_to = "Tasa"
  ) %>%
  mutate(
    Metrica = recode(Metrica, CAS = "Tasa Casos", PER = "Tasa Personas"),
    Sexo    = recode(Sexo, MUJ = "Mujeres", HOM = "Hombres"),
    Sexo    = factor(Sexo, levels = c("Hombres", "Mujeres")),
    Departamento = factor(Departamento)
  )

# ---------------------------------------------------------------------------
# 2. Estadísticos descriptivos: media, DE, mediana, CV, por sexo y métrica
# ---------------------------------------------------------------------------
cv <- function(x) (sd(x) / mean(x)) * 100

tabla3_desc <- suicidio_long %>%
  group_by(Metrica, Sexo) %>%
  summarise(
    Media   = mean(Tasa),
    DE      = sd(Tasa),
    Mediana = median(Tasa),
    CV      = cv(Tasa),
    .groups = "drop"
  )

print(tabla3_desc, n = Inf)

# ---------------------------------------------------------------------------
# 3. Modelo mixto: efecto del sexo sobre la Tasa de Casos
# ---------------------------------------------------------------------------
datos_casos <- suicidio_long %>% filter(Metrica == "Tasa Casos")

options(contrasts = c("contr.sum", "contr.poly"))

modelo_suicidio <- lmer(
  Tasa ~ Sexo + (1 | Departamento),
  data = datos_casos,
  REML = TRUE
)

anova(modelo_suicidio, type = 3)
eta_squared(modelo_suicidio, partial = TRUE)
r2(modelo_suicidio)

# Diferencia promedio Mujeres - Hombres (en tasa de casos por 100.000 hab.)
datos_casos %>%
  group_by(Sexo) %>%
  summarise(media = mean(Tasa)) %>%
  pivot_wider(names_from = Sexo, values_from = media) %>%
  mutate(Diferencia = Mujeres - Hombres)

# ---------------------------------------------------------------------------
# 4. Correlaciones de Spearman: Tasa de Personas con intento de suicidio
#    vs. cada grupo diagnóstico CIE-10, separado por sexo
# ---------------------------------------------------------------------------
grupos_dx <- c("F00-F09","F10-F19","F20-F29","F30-F39","F40-F48",
               "F50-F59","F60-F69","F70-F79","F80-F89","F90-F98","F99")

calcular_correlaciones <- function(sexo_sufijo) {
  var_suicidio <- TAS_SALMENTAL_2023[[paste0("TASA_INT_SUC_PER_", sexo_sufijo)]]
  
  map_dfr(grupos_dx, function(g) {
    var_dx <- TAS_SALMENTAL_2023[[paste0("TASA_", g, "_", sexo_sufijo)]]
    test <- suppressWarnings(cor.test(var_suicidio, var_dx, method = "spearman"))
    tibble(
      Grupo_diagnostico = g,
      rho    = unname(test$estimate),
      p      = test$p.value
    )
  })
}

correlaciones_mujeres <- calcular_correlaciones("MUJ") %>% mutate(Sexo = "Mujeres")
correlaciones_hombres <- calcular_correlaciones("HOM") %>% mutate(Sexo = "Hombres")

tabla_correlaciones <- bind_rows(correlaciones_mujeres, correlaciones_hombres) %>%
  mutate(sig = case_when(p < .001 ~ "***", p < .01 ~ "**", p < .05 ~ "*", TRUE ~ ""))

# Revisa aquí específicamente los 4 hallazgos que reportaste para mujeres:
tabla_correlaciones %>%
  filter(Sexo == "Mujeres", Grupo_diagnostico %in% c("F30-F39","F40-F48","F60-F69","F70-F79")) %>%
  print()

# Y confirma que en hombres ninguna es significativa:
tabla_correlaciones %>%
  filter(Sexo == "Hombres") %>%
  arrange(p) %>%
  print(n = Inf)

# ---------------------------------------------------------------------------
# 5. Formatear Tabla 3 mejorada en estilo APA y exportar a Word
# ---------------------------------------------------------------------------
formatear_p_apa <- function(p) {
  case_when(
    p < .001 ~ "< .001",
    TRUE     ~ sub("^0", "", sprintf("%.3f", p))
  )
}

tabla3_apa <- tabla3_desc %>%
  mutate(across(c(Media, DE, Mediana, CV), ~ sprintf("%.2f", .x))) %>%
  pivot_wider(
    names_from = Metrica,
    values_from = c(Media, DE, Mediana, CV),
    names_glue = "{Metrica}_{.value}"
  )

ft3 <- flextable(tabla3_apa) %>%
  set_table_properties(layout = "autofit") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  fontsize(size = 12, part = "all") %>%
  bold(part = "header") %>%
  align(align = "center", part = "all") %>%
  border_remove() %>%
  hline_top(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline_bottom(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline(i = 1, part = "header", border = fp_border(width = 1, color = "black")) %>%
  padding(padding = 4, part = "all")

# Tabla de correlaciones en formato APA (una fila por grupo dx x sexo)
tabla_corr_apa <- tabla_correlaciones %>%
  mutate(
    rho_fmt = sprintf("%.2f", rho),
    p_fmt   = paste0(formatear_p_apa(p), sig)
  ) %>%
  select(Grupo_diagnostico, Sexo, rho_fmt, p_fmt) %>%
  pivot_wider(names_from = Sexo, values_from = c(rho_fmt, p_fmt))

ft4 <- flextable(tabla_corr_apa) %>%
  set_table_properties(layout = "autofit") %>%
  font(fontname = "Times New Roman", part = "all") %>%
  fontsize(size = 12, part = "all") %>%
  bold(part = "header") %>%
  align(align = "center", part = "all") %>%
  border_remove() %>%
  hline_top(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline_bottom(part = "all", border = fp_border(width = 1.5, color = "black")) %>%
  hline(i = 1, part = "header", border = fp_border(width = 1, color = "black")) %>%
  padding(padding = 4, part = "all")

# ---------------------------------------------------------------------------
# 6. Documento Word con ambas tablas (APA 7)
# ---------------------------------------------------------------------------
numero_t3 <- fpar(ftext("Tabla 3", prop = fp_text(font.family = "Times New Roman", font.size = 12, bold = TRUE)))
titulo_t3 <- fpar(ftext("Tasas de intento de suicidio por sexo, Colombia, 2023",
                        prop = fp_text(font.family = "Times New Roman", font.size = 12, italic = TRUE)))
nota_t3 <- fpar(
  ftext("Nota. ", prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = TRUE)),
  ftext("Tasas por 100.000 habitantes. DE = desviación estándar; CV = coeficiente de variación (%).",
        prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = FALSE))
)

numero_t4 <- fpar(ftext("Tabla 4", prop = fp_text(font.family = "Times New Roman", font.size = 12, bold = TRUE)))
titulo_t4 <- fpar(ftext("Correlaciones de Spearman entre la tasa de intento de suicidio (personas) y las tasas diagnósticas, por sexo",
                        prop = fp_text(font.family = "Times New Roman", font.size = 12, italic = TRUE)))
nota_t4 <- fpar(
  ftext("Nota. ", prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = TRUE)),
  ftext("Rho = coeficiente de correlación de Spearman.",
        prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = FALSE))
)
nota_sig <- fpar(ftext("*p < .05. **p < .01. ***p < .001.",
                       prop = fp_text(font.family = "Times New Roman", font.size = 11, italic = TRUE)))

doc <- read_docx() %>%
  body_add_fpar(numero_t3) %>%
  body_add_fpar(titulo_t3) %>%
  body_add_flextable(ft3) %>%
  body_add_fpar(nota_t3) %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(numero_t4) %>%
  body_add_fpar(titulo_t4) %>%
  body_add_flextable(ft4) %>%
  body_add_fpar(nota_t4) %>%
  body_add_fpar(nota_sig)

print(doc, target = "Tabla3_4_APA.docx")

# ---------------------------------------------------------------------------
# 1. Reconstruir datos de tasa de CASOS de intento de suicidio (formato largo)
# ---------------------------------------------------------------------------
casos_suicidio_long <- TAS_SALMENTAL_2023 %>%
  select(Departamento, TASA_INT_SUC_CAS_MUJ, TASA_INT_SUC_CAS_HOM) %>%
  pivot_longer(
    cols = -Departamento,
    names_to  = "Sexo",
    names_prefix = "TASA_INT_SUC_CAS_",
    values_to = "TasaCasos"
  ) %>%
  mutate(
    Sexo = recode(Sexo, MUJ = "Mujeres", HOM = "Hombres"),
    Sexo = factor(Sexo, levels = c("Hombres", "Mujeres")),
    Departamento = factor(Departamento)
  )

# Debe quedar: 33 departamentos x 2 sexos = 66 filas
dim(casos_suicidio_long)
str(casos_suicidio_long)

# ---------------------------------------------------------------------------
# 2. Descriptivos por sexo (para comparar contra tu Tabla 3)
# ---------------------------------------------------------------------------
casos_suicidio_long %>%
  group_by(Sexo) %>%
  summarise(
    Media   = mean(TasaCasos),
    DE      = sd(TasaCasos),
    Mediana = median(TasaCasos),
    CV      = sd(TasaCasos) / mean(TasaCasos) * 100
  )

# ---------------------------------------------------------------------------
# 3. Modelo mixto: efecto del sexo sobre la tasa de casos
# ---------------------------------------------------------------------------
options(contrasts = c("contr.sum", "contr.poly"))

modelo_suicidio <- lmer(
  TasaCasos ~ Sexo + (1 | Departamento),
  data = casos_suicidio_long,
  REML = TRUE
)

# ANOVA tipo III (Satterthwaite) — esto es lo que da tu F(1,32), p, etc.
anova(modelo_suicidio, type = 3)

# Tamaño de efecto (eta parcial al cuadrado)
eta_squared(modelo_suicidio, partial = TRUE)

# R² marginal y condicional
r2(modelo_suicidio)

# Diferencia promedio Mujeres - Hombres en tasa de casos
casos_suicidio_long %>%
  group_by(Sexo) %>%
  summarise(media = mean(TasaCasos)) %>%
  pivot_wider(names_from = Sexo, values_from = media) %>%
  mutate(Diferencia_Mujeres_menos_Hombres = Mujeres - Hombres)

#figuras:
library(tidyverse)
library(scales)
library(tidytext)   # para reorder_within() / scale_y_reordered()

# ---------------------------------------------------------------------------
# 0. Formateador de decimales con COMA (para todas las figuras)
# ---------------------------------------------------------------------------
coma <- function(x, accuracy = 0.01) {
  label_number(accuracy = accuracy, decimal.mark = ",", big.mark = ".")(x)
}

# ---------------------------------------------------------------------------
# 1. Definir variables de demanda (por sexo) y de oferta
# ---------------------------------------------------------------------------
grupos_dx <- c("F00-F09","F10-F19","F20-F29","F30-F39","F40-F48",
               "F50-F59","F60-F69","F70-F79","F80-F89","F90-F98","F99")

etiquetas_dx <- c(
  "F00-F09" = "F00–F09 Orgánicos",
  "F10-F19" = "F10–F19 Sustancias",
  "F20-F29" = "F20–F29 Esquizofrenia y delirantes",
  "F30-F39" = "F30–F39 Estado de ánimo",
  "F40-F48" = "F40–F49 Neuróticos y estrés",
  "F50-F59" = "F50–F59 Conducta y fisiológicos",
  "F60-F69" = "F60–F69 Personalidad",
  "F70-F79" = "F70–F79 Retraso mental",
  "F80-F89" = "F80–F89 Desarrollo psicológico",
  "F90-F98" = "F90–F98 Infancia y adolescencia",
  "F99"     = "F99 No especificados"
)

vars_oferta <- c(
  "TASA_PSIQUI"       = "Psiquiatras, tasa",
  "TASA_PSICO"        = "Psicólogos, tasa",
  "TASA_Cama_SM"      = "Camas SM, tasa",
  "TASA_Camillas_SM"  = "Camillas SM, tasa",
  "TASA_Sillas_SM"    = "Sillas SM, tasa",
  "TASA_Camas_SPA"    = "Camas SPA, tasa",
  "TASA_Camillas_SPA" = "Camillas SPA, tasa",
  "TASA_Sillas_SPA"   = "Sillas SPA, tasa"
)

# ---------------------------------------------------------------------------
# 2. Función: matriz de correlaciones de Spearman (demanda x oferta) por sexo
# ---------------------------------------------------------------------------
matriz_correlacion <- function(sexo_sufijo) {
  
  # Variables de demanda: 11 grupos dx + intentos de suicidio (tasa personas)
  vars_demanda <- c(paste0("TASA_", grupos_dx, "_", sexo_sufijo),
                    paste0("TASA_INT_SUC_PER_", sexo_sufijo))
  etiquetas_demanda <- c(etiquetas_dx, "Intentos de suicidio, tasa")
  
  resultados <- expand_grid(
    demanda = vars_demanda,
    oferta  = names(vars_oferta)
  ) %>%
    mutate(
      rho = map2_dbl(demanda, oferta, ~ suppressWarnings(
        cor.test(TAS_SALMENTAL_2023[[.x]], TAS_SALMENTAL_2023[[.y]], method = "spearman")$estimate
      )),
      p = map2_dbl(demanda, oferta, ~ suppressWarnings(
        cor.test(TAS_SALMENTAL_2023[[.x]], TAS_SALMENTAL_2023[[.y]], method = "spearman")$p.value
      ))
    ) %>%
    mutate(
      demanda_lbl = factor(demanda, levels = vars_demanda, labels = etiquetas_demanda),
      oferta_lbl  = factor(oferta, levels = names(vars_oferta), labels = vars_oferta),
      sig = case_when(p < .001 ~ "***", p < .01 ~ "**", p < .05 ~ "*", TRUE ~ ""),
      etiqueta = paste0(coma(rho, 0.01), sig)
    )
  
  resultados
}

corr_mujeres <- matriz_correlacion("MUJ")
corr_hombres <- matriz_correlacion("HOM")

# ---------------------------------------------------------------------------
# 3. Función de graficado del heatmap (gris, decimales con coma)
# ---------------------------------------------------------------------------
graficar_heatmap <- function(df, titulo) {
  ggplot(df, aes(x = oferta_lbl, y = fct_rev(demanda_lbl), fill = rho)) +
    geom_tile(color = "white") +
    geom_text(aes(label = etiqueta,
                  fontface = ifelse(p < .05, "bold", "plain"),
                  color = abs(rho) > 0.5),
              size = 3.2, show.legend = FALSE) +
    scale_color_manual(values = c("black", "white")) +
    scale_fill_gradient(
      low = "white", high = "black", limits = c(-1, 1),
      labels = coma, name = "\u03c1 de Spearman,\nadimensional, rango -1 a 1"
    ) +
    labs(
      title = titulo,
      x = "Oferta, tasas por 100 000 habitantes",
      y = "Demanda: diagnósticos e intentos de suicidio, tasas por 100 000 habitantes"
    ) +
    theme_minimal(base_size = 11) +
    theme(
      axis.text.x = element_text(angle = 40, hjust = 1),
      panel.grid = element_blank(),
      plot.title = element_text(face = "bold")
    )
}

fig1_mujeres <- graficar_heatmap(corr_mujeres, "Mujeres: demanda de diagnósticos e intentos de suicidio vs oferta")
fig2_hombres <- graficar_heatmap(corr_hombres, "Hombres: demanda de diagnósticos e intentos de suicidio vs oferta")

ggsave("Figura1_mujeres_oferta_demanda.png", fig1_mujeres, width = 10, height = 7, dpi = 300)
ggsave("Figura2_hombres_oferta_demanda.png", fig2_hombres, width = 10, height = 7, dpi = 300)

# ---------------------------------------------------------------------------
# 4. Panel de barras horizontales: 33 departamentos x grupo diagnóstico CIE-10
#    (bidimensional: un panel por diagnóstico, ordenado por tasa)
#    NOTA: como no hay tasa poblacional total por departamento, la "tasa total"
#    se calcula como el promedio simple de la tasa de mujeres y hombres.
#    Si tienes la tasa total real (ponderada por población), reemplaza esto.
# ---------------------------------------------------------------------------
tasas_totales_dx <- map_dfr(grupos_dx, function(g) {
  tibble(
    Departamento = TAS_SALMENTAL_2023$Departamento,
    Diagnostico  = etiquetas_dx[[g]],
    Tasa_total   = (TAS_SALMENTAL_2023[[paste0("TASA_", g, "_MUJ")]] +
                      TAS_SALMENTAL_2023[[paste0("TASA_", g, "_HOM")]]) / 2
  )
})

fig_barras <- tasas_totales_dx %>%
  mutate(Departamento_ord = reorder_within(Departamento, Tasa_total, Diagnostico)) %>%
  ggplot(aes(x = Tasa_total, y = Departamento_ord)) +
  geom_col(fill = "grey30") +
  facet_wrap(~ Diagnostico, scales = "free", ncol = 3) +
  scale_y_reordered() +
  scale_x_continuous(labels = coma) +
  labs(
    title = "Tasas por grupo diagnóstico CIE-10 según departamento, Colombia, 2023",
    x = "Tasa por 100 000 habitantes (promedio mujeres-hombres)",
    y = NULL
  ) +
  theme_minimal(base_size = 8) +
  theme(
    strip.text = element_text(face = "bold", size = 7),
    axis.text.y = element_text(size = 5),
    panel.grid.minor = element_blank()
  )

ggsave("Figura3_barras_departamentos_CIE10.png", fig_barras, width = 14, height = 16, dpi = 300)