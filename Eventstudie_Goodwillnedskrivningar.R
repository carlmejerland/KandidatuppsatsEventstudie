# ==========================================
# Kandidatarbete i externredovisning
# Författare: Simon Brandhammar & Car Mejerland
# Titel: Marknadsreaktionen på Goodwillnedskrivningar
# Datum: [2025-06-04]
# Fil: Kandidat_Kod
# Syfte: Eventstudie för att mäta CAAR vid goodwillnedskrivningar
# ==========================================

# Rensa miljön
rm(list = ls())
graphics.off()

# Bibliotek
library(readxl)
library(tidyverse)
library(dplyr)
library(glue)
library(ggplot2)
library(tidyr)
library(knitr)

# ===== Funktioner =====

# 1) Läs in rådata
las_in_data <- function(fil) {
  aktier_raw    <- read_excel(fil, sheet = "aktiekurser", col_types = "text")
  index_raw     <- read_excel(fil, sheet = "index",      col_types = "text")
  eventdata_raw <- read_excel(
    fil,
    sheet     = "event",
    col_types = c("text", "text", "date", "text", "numeric", "numeric")
  )
  controls_raw   <- read_excel("Data.xlsx", sheet = "controls")
  list(
    aktier_raw    = aktier_raw,
    index_raw     = index_raw,
    eventdata_raw = eventdata_raw,
    controls_raw  = controls_raw
  )
}

# 2) Formatera aktie- och indexdata till return_list
formatera_return_list <- function(aktier_raw, index_raw) {
  # Aktiedata
  aktier_list <- aktier_raw %>%
    pivot_longer(-Företag, names_to = "Date", values_to = "Close_raw") %>%
    mutate(
      Date  = as.Date(as.numeric(Date), origin = "1899-12-30"),
      Close = as.numeric(str_replace(Close_raw, ",", "."))
    ) %>%
    filter(!is.na(Close), Close > 0) %>%
    group_by(Företag) %>%
    arrange(Date) %>%
    mutate(
      Return    = log(Close / lag(Close))
    ) %>%
    ungroup()
  
  # Indexdata
  index_list <- index_raw %>%
    pivot_longer(-Datum, names_to = "Date", values_to = "Marknadsindex") %>%
    mutate(
      Date           = as.Date(as.numeric(Date), origin = "1899-12-30"),
      Marknadsindex  = as.numeric(Marknadsindex),
      Marknadsreturn = log(Marknadsindex / lag(Marknadsindex))
    ) %>%
    select(-Datum)
  
  # Slå ihop
  left_join(aktier_list, index_list, by = "Date") %>%
    select(Företag, Date, Return, Marknadsreturn) %>%
    drop_na(Return, Marknadsreturn) %>%
  group_by(Företag) %>%
    arrange(Date) %>%
    mutate(
      Radnummer = row_number()
    ) %>%
    ungroup()
}

# 3) Skapa eventlista
skapa_eventlista <- function(eventdata_raw) {
  lapply(seq_len(nrow(eventdata_raw)), function(i) {
    rad <- eventdata_raw[i, ]
    list(
      Företag      = rad$Företag,
      Datum        = as.Date(rad$Datum),
      Nedskrivning = rad$Nedskrivning,
      Nivå         = rad$Nivå,
      Alpha        = NA_real_,
      Beta         = NA_real_,
      Sigma2       = NA_real_,
      CAR          = NA_real_
    )
  })
}

# Rensa Controlls
rensa_controls <- function(controls_raw) {
  controls_raw %>%
    mutate(Datum = as.Date(Datum)) %>%
    select(Företag, Datum, MarketCap, ROE, ROA) %>%
    mutate(
      MarketCap = as.numeric(str_replace(MarketCap, ",", ".")),
      #TV        = as.numeric(str_replace(TV, ",", ".")),
      ROE       = as.numeric(str_replace(ROE, ",", ".")) / 100,
      ROA       = as.numeric(str_replace(ROA, ",", ".")) / 100
    )
}

# 4) Skapa estimeringslista (ta bort eventfönster)
create_estimation_list <- function(return_list, events, est_gap) {
  est_list <- return_list
  for (ev in events) {
    rad <- est_list %>%
      filter(Företag == ev$Företag, Date <= ev$Datum) %>%
      arrange(desc(Date)) %>%
      slice(1) %>%
      pull(Radnummer)
    if (length(rad) == 0) next
    est_list <- est_list %>%
      filter(
        !(Företag == ev$Företag &
            Radnummer >= (rad - est_gap) &
            Radnummer <= (rad + est_gap))
      )
  }
  est_list
}

# 5) Skatta α, β och σ² för ett event
skatta_marknadsmodell <- function(est_list, ev, est_win, est_gap) {
  rad <- est_list %>%
    filter(Företag == ev$Företag, Date <= ev$Datum) %>%
    arrange(desc(Date)) %>%
    slice(1) %>%
    pull(Radnummer)
  if (length(rad) != 1) {
    return(list(Alpha = NA_real_, Beta = NA_real_, Sigma2 = NA_real_))
  }
  df <- est_list %>%
    filter(
      Företag   == ev$Företag,
      Radnummer >= rad - est_win,
      Radnummer <= rad - est_gap - 1
    ) %>%
    filter(if_all(c(Return, Marknadsreturn), is.finite))
  m  <- lm(Return ~ Marknadsreturn, data = df)
  co <- coef(m)
  s2 <- summary(m)$sigma^2
  list(Alpha = co["(Intercept)"], Beta = co["Marknadsreturn"], Sigma2 = s2)
}

# 6) Beräkna AR
beräkna_AR <- function(rad, alpha, beta) {
  rad$Return - (alpha + beta * rad$Marknadsreturn)
}

# 7) Beräkna CAR
beräkna_CAR <- function(return_lista, ref_datum, pre, post, alpha, beta) {
  entry <- return_lista %>%
    filter(Date >= ref_datum) %>%
    arrange(Date) %>%
    slice(1)
  if (nrow(entry) != 1) {
    stop(glue::glue("beräkna_CAR: ingen rad med Date ≥ {ref_datum}"))
  }
  used_date <- entry$Date
  r         <- entry$Radnummer
  if (used_date > ref_datum) {
    warning(glue::glue("beräkna_CAR: använder nästa handelsdag {used_date}"))
  }
  win <- return_lista %>%
    filter(
      Radnummer >= r + pre,
      Radnummer <= r + post
    ) %>%
    filter(if_all(c(Return, Marknadsreturn), is.finite))
  if (nrow(win) == 0) {
    warning(glue::glue("beräkna_CAR: inga giltiga rader för datum {used_date}"))
    return(NA_real_)
  }
  sum(map_dbl(seq_len(nrow(win)), ~ beräkna_AR(win[.x, ], alpha, beta)), na.rm = TRUE)
}

# 8) Sammanfatta CAAR över alla events
sammanfatta_CAAR <- function(ev_list, pre, post) {
  CARs <- map_dbl(ev_list, "CAR")
  S2   <- map_dbl(ev_list, "Sigma2")
  CARn <- CARs[!is.na(CARs)]
  S2n  <- S2[!is.na(CARs)]
  N    <- length(CARn)
  wlen <- abs(post - pre) + 1
  var  <- sum(wlen * S2n) / (N^2)
  se   <- sqrt(var)
  CAAR <- mean(CARn)
  t    <- CAAR / se
  p    <- 2 * pt(-abs(t), df = N - 1)
  list(N = N, CAAR = CAAR, SE = se, t = t, p = p)
}


# Ligger i krisperiod
ligger_i_krisperiod <- function(datum, perioder) {
  any(sapply(perioder, function(p) datum >= p[1] & datum <= p[2]))
}



# ===== Parametrar =====
fil                <- "Data.xlsx"
est_gap            <- 10
est_win            <- 130
event_window_pre   <- 0
event_window_post  <- 2
trim_proportion    <- 0   # Sätt till 0 om du inte vill ta bort något, annars hur mycket ska tas bort från
finanskris_perioder <- list(
  as.Date(c("2008-01-01", "2009-12-31")),
  as.Date(c("2020-01-01", "2021-12-31"))
)


# ===== Körning =====

# Skapa listor
data_raw       <- las_in_data(fil)
return_list    <- formatera_return_list(data_raw$aktier_raw, data_raw$index_raw)
events         <- skapa_eventlista(data_raw$eventdata_raw)
controls_clean <- rensa_controls(data_raw$controls_raw)
est_list       <- create_estimation_list(
  return_list, events,
  est_gap
)

# Marknadsmodell för alla event
for (i in seq_along(events)) {
  m <- skatta_marknadsmodell(
    est_list = est_list,
    ev       = events[[i]],
    est_win  = est_win,
    est_gap  = est_gap
  )
  events[[i]]$Alpha  <- m$Alpha
  events[[i]]$Beta   <- m$Beta
  events[[i]]$Sigma2 <- m$Sigma2
}

# CAR för alla event
for (i in seq_along(events)) {
  rl <- return_list %>%
    filter(Företag == events[[i]]$Företag)
  events[[i]]$CAR <- beräkna_CAR(
    return_lista = rl,
    ref_datum    = events[[i]]$Datum,
    pre          = event_window_pre,
    post         = event_window_post,
    alpha        = events[[i]]$Alpha,
    beta         = events[[i]]$Beta
  )
}

#Trimma extremvärden
if (trim_proportion > 0) {
  car_vals <- map_dbl(events, "CAR")
  lower    <- quantile(car_vals, trim_proportion, na.rm = TRUE)
  upper    <- quantile(car_vals, 1 - trim_proportion, na.rm = TRUE)
  events   <- events[car_vals >= lower & car_vals <= upper]
}

# Beräkna CAAR
res <- sammanfatta_CAAR(events, event_window_pre, event_window_post)
print(res)


# Bygg car_df
car_df <- tibble(
  Företag      = map_chr(events, "Företag"),
  #Nivå         = map_dbl(events, "Nivå"),
  #Nedskrivning = map_dbl(events, "Nedskrivning"),
  Datum        = map(events, "Datum")   %>% flatten_dbl() %>% as_date(),
  CAR          = map_dbl(events, "CAR")
) %>%
  left_join(controls_clean, by = c("Företag", "Datum")) %>%
  mutate(
    #ImpairmentRatio = Nedskrivning / TV,
    logMarketCap    = log(MarketCap),
    FinanskrisDummy      = if_else(
      map_lgl(Datum, ~ ligger_i_krisperiod(.x, finanskris_perioder)),
      1L, 0L
    )
  )

# Multivariat regression
modell <- lm(CAR ~  ROE + ROA + logMarketCap + FinanskrisDummy, data = car_df)
print(summary(modell))

# Deskriptiv statistik
stats <- car_df %>%
  summarise(
    mean_CAR   = mean(CAR,   na.rm = TRUE),
    sd_CAR     = sd(CAR,     na.rm = TRUE),
    var_CAR    = var(CAR,    na.rm = TRUE),
    median_CAR = median(CAR, na.rm = TRUE),
    min_CAR    = min(CAR,    na.rm = TRUE),
    max_CAR    = max(CAR,    na.rm = TRUE)
  )
print(stats)


car_df <- car_df %>%
  arrange(CAR) %>%
  mutate(rownum = row_number())

ggplot(car_df, aes(x = rownum, y = CAR, fill = rownum)) +
  geom_col(show.legend = FALSE) +
  labs(title = "", x = "Event", y = "CAR") +
  scale_x_continuous(breaks = c(1, 44, 87), labels = c("1", "44", "87")) +
  theme_minimal()

