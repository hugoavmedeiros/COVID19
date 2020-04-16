  library(ggplot2)
  library(openxlsx)
  library(officer)
  library(rvg)
  library(usmap)
  library(httr)
  library(cowplot)
  library(dplyr)
  library(bfw)
  library(rnaturalearth)
  library(sf)
  
  remove(list=ls())
  setwd("g:/projects/covid")
  today <- Sys.Date()  
  startDate <- as.Date("2020-01-22")
  allDays <- as.numeric(today-startDate)
  
  source("ImportCovidData.R")
  # USA Files: Cases_USA, Deaths_USA, Population_USA
  # Global Files: Cases_Global, Deaths_Global, Population_Global
  
  lastColumn <- which(
    grepl(gsub("^X0","X", format(today - 1, "X%m.%e.%y")), names(Cases_USA),1) |
    grepl(gsub("^X0","X", format(today - 1, "X%m.%e.%Y")), names(Cases_USA),1)
  )

if (length(lastColumn) > 0)
{
  Cases_USA  <- Cases_USA[,1:lastColumn]
  Deaths_USA <- Deaths_USA[,1:lastColumn]
  
  projection <- 7 # Project forward just 7 days
  asymptomatic <- 10  # Number of asymptomatic patients per symptomatic patient
  endDate <- startDate + allDays + projection - 1
  Dates <- seq(from = startDate, to=endDate, by = 1)
  extendedDates <- seq(from = startDate, to=startDate + allDays + 60, by = 1)
  bootstrapsN <- 100

  timestamp <- format(Sys.time(), format = "%Y-%m-%d")
  pptx <- read_pptx("Template.pptx")
  master <- "Office Theme"
  slideNumber <- 1
  pptxfileName <- paste0("Steve's COVID Analysis.", timestamp, ".pptx")
  if (file.exists(pptxfileName))
    file.remove(pptxfileName)
  
  while(file.exists(pptxfileName))
  {
    cat("Close the open PowerPoint File\n")
    file.remove(pptxfileName)
  }
  
  
  # Disclaimer slide
  DISCLAIMER <-
    c(
      "This is not confidential and can be freely shared. The code is available at https://github.com/StevenLShafer/COVID19/.",
      "",
      "This is my analysis, not Stanford's.",
      "",
      "Data sources:",
      "       USA Data:     https://usafactsstatic.blob.core.windows.net/public/data/covid-19/covid_confirmed_usafacts.csv.",
      "       Global Data:  https://github.com/CSSEGISandData/COVID-19. This is the Johns Hopkins data repository.",
      "",
      "Nate Silver has an excellent write-up on the potential problems of the data (see https://fivethirtyeight.com/features/coronavirus-case-counts-are-meaningless/). He is right, of course, but these data are all we have. Also, he does not address the fact that the data are consistent with the expections. Specifically, the increases are initially log-linear, but then \"flatten\" as expected when public health policies are implemented. We would not see this if the data were just random noise.",
      "",
      "Models:",
      "       Primary (applied to the red dots on the graph): log(y) = intercept + (peak - intercept) * (1 - exp(k * time))",
      "       Log linear: log(Y) = intercept + slope * time, used to compute doubling time (0.693/slope).",
      "",
      "The number printed on the graph is the projection for a week from today. Doubling times are calculated for 5 day windows.",
      "",
      "The idiosyncratic locations are where Pamela and I have family or friends, or are locations requested by friends. I'm happy to add other regions. Also, I'm happy to add people to the blind CC distribution list. Just let me know.",
      "",
      "Please send any questions to steven.shafer@stanford.edu.",
      "",
      "Stay safe, well, and kind."
    )

  pptx <- add_slide(pptx, layout = "Title and Text", master = master)
  pptx <- ph_with(pptx, value = "Caveats and Comments",  location = ph_location_type("title") )
  pptx <- ph_with(pptx, value = "", location = ph_location_type("body"))
  pptx <- ph_add_text(pptx, str = paste(DISCLAIMER, collapse = "\n"), style = fp_text(font.size = 10))
  pptx <- ph_with(pptx, value = slideNumber, location = ph_location_type("sldNum"))
  slideNumber <- slideNumber + 1

  nextSlide <- function (ggObject, Title)
    {
    suppressWarnings(
      print(ggObject)
    )

    pptx <- add_slide(pptx, layout = "Title and Content", master = master)
    pptx <- ph_with(pptx, value = Title, location = ph_location_type("title"))
    suppressWarnings(
      pptx <<- ph_with(pptx, value = dml(ggobj = ggObject), location = ph_location_type("body"))
    )
    pptx <<- ph_with(pptx, value = slideNumber, location = ph_location_type("sldNum"))
    slideNumber <<- slideNumber + 1
  }
  
closest <- function(a, b)
{
  which(abs(a-b) == min(abs(a-b), na.rm=TRUE))[1]
}

covid_fn <- function(par, N)
{
  # par <- intercept, peak, k
  return(par[1] + (par[2] - par[1]) * (1-exp(-par[3] * 0:(N-1))))
}

covid_obj <- function(par, Y, weight)
{
  return(
    sum(
      (Y-covid_fn(par, length(Y)))^2 * (1:length(Y))^weight
  )
  )
}

covid_fit <-  function(Y, maxCases, weight = 1)
{
  Y1 <- tail(Y,5)
  X1 <- 1:5
  slope <- lm(Y1 ~ X1)$coefficients[2]
  return(
    optim(
      c(Y[1],Y[1] + 2, slope),
      covid_obj,
      Y = Y,
      weight = weight,
    method = "L-BFGS-B",
    lower = c(Y[1] - 1, Y[1] - 1, 0.01),
    upper = c(Y[1] + 1, maxCases, 0.693)
    )$par
  )
}


plotPred <- function(
  County = NULL, 
  State = NULL, 
  Country = NULL,
  Title, 
  logStart, 
  logEnd,
  weight = 1
  )
  {
  
  # USA Data
  if (!is.null(County) | !is.null(State))
  {
    if(is.null(County))
    {
      useCounty <- rep(TRUE, nrow(Cases_USA))
    } else {
      useCounty <- Cases_USA$County.Name %in% County
    }
    if(is.null(State))
    {
      useState <- rep(TRUE, nrow(Cases_USA))
    } else {
      useState <- Cases_USA$State %in% State
    }
    
    use <- useCounty & useState

    CASES <- data.frame(
      Date = Dates,
      Actual = c(colSums(Cases_USA[use, c(5:ncol(Cases_USA))],na.rm=TRUE), rep(NA, projection)),
      Phase = "",
      Predicted = NA,
      stringsAsFactors = FALSE
    )
    
    DEATHS <- data.frame(
      Date = Dates,
      Actual = c(colSums(Deaths_USA[use, c(5:ncol(Deaths_USA))],na.rm=TRUE), rep(NA, projection)),
      Phase = "Deaths",
      Predicted = NA,
      stringsAsFactors = FALSE
    )
    FIPS <- Cases_USA$CountyFIPS[use]
    maxCases <- sum(Population_USA$Population[Population_USA$CountyFIPS %in% FIPS], na.rm=TRUE) / (1 + asymptomatic)
  }
  
  # Global data
  if (!is.null(Country))
  {
    use <- Cases_Global$Country %in% Country
    CASES <- data.frame(
      Date = Dates,
      Actual = c(colSums(Cases_Global[use, c(2:ncol(Cases_Global))],na.rm=TRUE), rep(NA, projection)),
      Phase = "",
      Predicted = NA,
      stringsAsFactors = FALSE
    )
    
    DEATHS <- data.frame(
      Date = Dates,
      Actual = c(colSums(Deaths_Global[use, c(2:ncol(Deaths_Global))],na.rm=TRUE), rep(NA, projection)),
      Phase = "Deaths",
      Predicted = NA,
      stringsAsFactors = FALSE
    )
    maxCases <- sum(Population_Global$Population[Population_Global$Country %in% Country], na.rm=TRUE) / (1 + asymptomatic)
  }
  
  # Cumulative case numbers and deaths cannot drop
  for (i in nrow(CASES):2)
  {
    if(!is.na(CASES$Actual[i]))
      if (CASES$Actual[i-1] > CASES$Actual[i]) CASES$Actual[i-1] <- CASES$Actual[i]
      if(!is.na(DEATHS$Actual[i]))
        if (DEATHS$Actual[i-1] > DEATHS$Actual[i]) DEATHS$Actual[i-1] <- DEATHS$Actual[i]
  }

    start <- which(CASES$Date == as.Date(logStart))
    end   <-  which(CASES$Date == as.Date(logEnd))
    CASES$Phase[1:(start-1)] <- "Pre log linear"
    CASES$Phase[start:(end-1)] <- "Log linear"
    CASES$Phase[end:nrow(CASES)] <- "Current"
    CASES$Actual[CASES$Actual == 0 | is.na(CASES$Actual)] <- NA
    DEATHS$Actual[DEATHS$Actual == 0 | is.na(DEATHS$Actual)] <- NA
    maxCases <- log(maxCases)

    # Current
    use <- CASES$Date >= as.Date(logEnd) & !is.na(CASES$Actual)
    Y <- log(CASES$Actual[use])
    fit <- covid_fit(Y, maxCases, weight)
    
    # Last 5 days
    Y <- log(CASES$Actual[CASES$Date > today - 6 & CASES$Date < today])
    if (Y[1] < Y[5])
    {
      X <- 1:length(Y)
      coefs <- lm(Y ~ X)$coefficients
      last5Doubling <- sprintf("%0.1f", 0.693 / coefs[2]) 
    } else {
      last5Doubling <- "flat (no change)"
    }

    # Prediction over next week
    X <- end:nrow(CASES)
    L <- length(X)
    CASES$Predicted[X] <- exp(covid_fn(fit, L))
    WeekPrediction <- tail(CASES$Predicted, 1)
    if (sum(DEATHS$Actual, na.rm = TRUE) > 0)
    {
      DATA <- rbind(CASES, DEATHS)
    } else {
      DATA <- CASES
    }

    caption <- paste0("Doubling time over past 5 days: ", last5Doubling) 

    DATA$Phase <- factor(as.character(DATA$Phase), levels=c("Pre log linear","Log linear","Current", "Deaths"), ordered = TRUE)
  #  DEATHS$Phase <- factor(as.character(DEATHS$Phase), levels=c("Pre log linear","Log linear","Current", "Deaths"), ordered = TRUE)
    ggObject <-
      ggplot(DATA, aes(x = Date, y = Actual, color = Phase)) +
      geom_point(size = 2, na.rm = TRUE) +
      coord_cartesian(ylim = c(1,10000000), expand = TRUE, clip = "on") +
      geom_line(data=DATA[DATA$Phase != "Deaths",], size = 1, aes(y=Predicted), color="red", na.rm = TRUE) +
      labs(
        title = paste(Title,"projection as of", Sys.Date()),
        y = "Actual (points) / Predicted (line)",
        caption = caption
      ) +
      scale_color_manual(values = c("forestgreen","blue","red", "black")) +
      scale_x_date(
        date_breaks = "7 days",
        date_labels = "%b %d"
      ) +
      scale_y_log10(
        breaks = c(1, 10, 100, 1000, 10000, 100000, 1000000,10000000),
        labels = c("1", "10","100","1,000","10,000","100,000","1,000,000", "10,000,000")
      ) +
      theme(axis.text.x=element_text(angle=60, hjust=1))+
      annotation_logticks() +
      theme(
        panel.grid.minor = element_blank(),
        legend.key = element_rect(fill = NA)
      ) +
      annotate(
        geom = "point",
        x = endDate,
        y = WeekPrediction,
        shape = 8
      ) +
      annotate(
        geom = "text",
        x = endDate,
        y = WeekPrediction * 1.4,
        hjust = 0.5,
        vjust = 0,
        label = prettyNum(round(WeekPrediction, 0), big.mark = ",", scientific = FALSE),
        size = 3
      )

    INSET <- data.frame(
      Date = extendedDates,
      Total = 0,
      Delta = 0,
      doublingTime = NA, 
      stringsAsFactors = FALSE
    )
   
    for (i in 2:nrow(CASES))
    {
      INSET$Total[i] <- CASES$Actual[i]
    }
    X <- end:nrow(INSET)
    L <- length(X)
    INSET$Total[X] <- exp(covid_fn(fit, L))

    # Cumulative case numbers cannot drop, needed because of merging reported cases with predicted cases.
    N <- nrow(INSET)
    INSET$Total[is.na(INSET$Total)] <- 0
    for (i in N:2)
    {
      if (INSET$Total[i-1] > INSET$Total[i]) INSET$Total[i-1] <- INSET$Total[i]
    }
    INSET$Total[is.na(INSET$Total)] <- 0
    
    # Calculate Doubling Times
    X <- 1:5
    for (i in 40:N)
    {
      if (INSET$Total[i-4] > 0)
      {
        Y <- log(INSET$Total[(i-4):i])
        INSET$doublingTime[i] <- 0.693 / lm(Y ~ X)$coefficients[2]
      }
    }
    INSET$doublingTime[INSET$doublingTime > 30] <- 30.123456 # No point in showing a doubling time > 30 days. This is a unique placeholder so other calculation can occur
    INSET$doublingTime[INSET$doublingTime < 0] <- 0

    INSET$Delta[2:N] <- INSET$Total[2:N] - INSET$Total[1:(N-1)]
    INSET$Source <- "Reported"
    INSET$Source[INSET$Date > today] <- "Predicted"
    
    Reported <- closest(INSET$Delta[INSET$Date <= today], max(INSET$Delta[INSET$Date <= today])/2)
    Predicted <- closest(INSET$Delta[INSET$Date > today], (max(INSET$Delta[INSET$Date > today])+min(INSET$Delta[INSET$Date > today]))/2) + allDays
    
    topINSET <- cbind(INSET[,c("Date", "Delta","Source")], "New Cases / Day")
    bottomINSET <- cbind(INSET[,c("Date", "doublingTime","Source")],"Doubling Time")
    bottomINSET$Source[INSET$Source == "Reported"] <- "Calculated"
    names(topINSET) <- names(bottomINSET) <- c("Date","Y","Source","Wrap")
    
    Reported <- 
      closest(
        topINSET$Y[topINSET$Date <= today], 
        max(topINSET$Y[topINSET$Date <= today], na.rm=TRUE)/2)
    Calculated <- 
      closest(
        bottomINSET$Y[bottomINSET$Date <= today], 
        max(bottomINSET$Y[bottomINSET$Date <= today], na.rm=TRUE)/2)+ nrow(topINSET)
    Predicted1 <- 
      closest(
        topINSET$Y[topINSET$Date > today], 
        (
          max(topINSET$Y[topINSET$Date > today], na.rm=TRUE)+
          min(topINSET$Y[topINSET$Date > today], na.rm=TRUE)
          )/2) + allDays 
    Predicted2 <- 
      closest(
        bottomINSET$Y[bottomINSET$Date > today], 
        (
          max(bottomINSET$Y[bottomINSET$Date > today], na.rm=TRUE)+
            min(bottomINSET$Y[bottomINSET$Date > today], na.rm=TRUE)
        )/2) + allDays + nrow(topINSET)
    
     INSET <- rbind(topINSET, bottomINSET)

    ann_text <- data.frame(
      Date = c(INSET$Date[Reported] - 1, INSET$Date[Predicted1] +1, INSET$Date[Calculated] - 1, INSET$Date[Predicted2] +1),
      Y =    c(INSET$Y[Reported], INSET$Y[Predicted1], INSET$Y[Calculated],INSET$Y[Predicted2]),
      Text = c("Reported","Predicted", "Calculated", "Predicted"),
      hjust = c(1.2, -0.2, 1.2, -0.2),
      vjust = c(-0.2, -0.2, -0.2, -0.2),
      color = c("black","brown","black","brown"),
      Wrap = c("New Cases / Day", "New Cases / Day", "Doubling Time","Doubling Time"),
      stringsAsFactors = FALSE
    )
    today_line <- data.frame(
      Date = c(today + 0.5, today + 0.5, today + 0.5, today + 0.5),
      Y = c(0, max(INSET$Y, na.rm = TRUE), 0, 30),
      Wrap = c("New Cases / Day", "New Cases / Day", "Doubling Time","Doubling Time")
    )
    
    today_text <- data.frame(
      Date = c(today, today), 
      Y =    c(0, 0), 
      Text = c("today","today"),
      hjust = c(-0.5, -0.5),
      vjust = c(-0.5, -0.5),
      Wrap = c("New Cases / Day", "Doubling Time"),
      stringsAsFactors = FALSE
    )
      
    if (topINSET$Y[topINSET$Date == today] < max(topINSET$Y)/2)
    {
      today_text$Y[1] <- max(topINSET$Y)
      today_text$hjust[1] <- 1.5
    }

    if (bottomINSET$Y[bottomINSET$Date == today] < 15)
    {
      today_text$Y[2] <- 30
      today_text$hjust[2] <- 1.5
    }
    
    INSET$Wrap <- factor(INSET$Wrap, c("New Cases / Day", "Doubling Time"), ordered = TRUE)
    ann_text$Wrap <- factor(ann_text$Wrap, c("New Cases / Day", "Doubling Time"), ordered = TRUE)
    today_line$Wrap <- factor(today_line$Wrap, c("New Cases / Day", "Doubling Time"), ordered = TRUE)
    today_text$Wrap <- factor(today_text$Wrap, c("New Cases / Day", "Doubling Time"), ordered = TRUE)
    INSET$Y[INSET$Y == 30.123456] <- NA  # Get rid of values > 30
    
  ggObject2 <- 
    ggplot(INSET, aes(x=Date, y = Y)) + 
    geom_point(data = INSET[INSET$Date <= today,], size = 0.7, na.rm = TRUE, show.legend = FALSE, color = "black") +
    geom_line(data = INSET[INSET$Date > today,], size = 0.8, na.rm = TRUE, show.legend = FALSE, color = "brown") +
    scale_color_manual(values = c("black","brown")) +
    labs(
      y = "New Cases / Day"
    ) +
    labs(y=NULL) +
    scale_y_continuous(
      labels = scales::comma_format(big.mark = ',',
                                  decimal.mark = '.')) +
    geom_text(data = ann_text, aes(label = Text, hjust = hjust, vjust = vjust, color = color), size = 2.5, show.legend = FALSE) +
    geom_text(data = today_text, aes(label = Text, hjust = hjust, vjust = vjust), color="blue", angle = 90, size=2, show.legend = FALSE) +
    geom_line(data = today_line, color = "blue") +
    theme(
      axis.text=element_text(size=6),
      axis.title=element_text(size=8),
      strip.background = element_blank(),
      strip.placement = "outside" 
      ) + 
    facet_grid(
      Wrap ~ .,
      scales="free_y",
      switch = "y",
      shrink=FALSE
  )
  
  ggObject3 <-  ggdraw() +
    draw_plot(ggObject) +
    draw_plot(ggObject2, x = 0.12, y = .405, width = .29, height = .50)
  nextSlide(ggObject3, Title)
}

  # US Data
  Country <- "United States of America"
  County <- NULL
  State <- NULL
  logStart <- "2020-02-29"
  logEnd   <- "2020-03-22"
  Title <- "USA"
  weight <- 1
  plotPred(Country = "United States of America", Title = "USA", logStart = "2020-02-28", logEnd = "2020-03-21")
  plotPred(County = "New York County", Title = "New York City", logStart = "2020-03-02", logEnd = "2020-03-21", 
           weight = 2)
  plotPred(State = "CA", Title = "California", logStart = "2020-03-02", logEnd = "2020-03-25", weight = 0)
  plotPred(County = c("Santa Clara County", "San Mateo County"), Title = "Santa Clara and San Mateo", 
           logStart = "2020-03-02", logEnd = "2020-03-20", weight = 1)
  plotPred(County = "San Francisco County", Title = "San Francisco", logStart = "2020-03-07", logEnd = "2020-03-27")
  plotPred(County = "San Luis Obispo County", Title = "San Luis Obispo", 
           logStart = "2020-03-16", logEnd = "2020-03-28")
  plotPred(County = "King County", State = "WA", Title = "King County (Seattle)", 
           logStart = "2020-02-29", logEnd = "2020-03-19", weight=0)
  plotPred(County = "Los Angeles County", State = "CA", Title = "Los Angeles", logStart = "2020-03-04", 
           logEnd = "2020-03-27", weight = 1.5)
  plotPred(County = "Multnomah County", Title = "Multnomah County (Portland)", 
           logStart = "2020-03-16", logEnd = "2020-03-31")
  plotPred(County = "Westchester County", Title = "Westchester County", logStart = "2020-03-15", 
           logEnd = "2020-03-26", weight = 1)
  plotPred(County = "Alameda County", Title = "Alameda County", logStart = "2020-03-05", logEnd = "2020-03-24")
  plotPred(County = c("Santa Clara County", "San Mateo County", "San Francisco County", "Marin County", "Napa County", "Solano County", "Sonoma County"), 
           Title = "Bay Area", logStart = "2020-03-02", logEnd = "2020-03-23")
  plotPred(County = "De Soto Parish", Title = "De Soto Parish, Louisiana", 
           logStart = "2020-03-22", logEnd = "2020-03-25")
  plotPred(County = "Bergen County", Title = "Bergen County", logStart = "2020-03-14", 
           logEnd = "2020-03-29", weight = 1)
  plotPred(State = "DC", Title = "Washington DC", logStart = "2020-03-14", 
           logEnd = "2020-04-03", weight = 1.5)
  plotPred(County = "Dallas County", State = "TX", Title = "Dallas Texas", 
           logStart = "2020-03-10", logEnd = "2020-03-25")
  plotPred(County = "Collin County", State = "TX", Title = "Collin Texas", 
           logStart = "2020-03-19", logEnd = "2020-03-26")
  plotPred(County = "Harris County", State = "TX", Title = "Harris County, Texas", 
           logStart = "2020-03-20", logEnd = "2020-04-07",weight = 1)
  plotPred(County = "McLean County", State = "IL", Title = "McLean County, Illinois", logStart = "2020-03-20",
           logEnd = "2020-04-02")
  plotPred(County = "Cook County", State = "IL", Title = "Cook County, Illinois", logStart = "2020-03-06", logEnd = "2020-03-20")
  plotPred(County = "Suffolk County", State = "MA", Title = "Suffolk County (Boston)", logStart = "2020-03-10", 
           logEnd = "2020-04-05", weight = 2)
  plotPred(State = "UT", Title = "Utah (State)", logStart = "2020-03-02", logEnd = "2020-03-18")
  plotPred(County = "Utah County", Title = "Utah County", logStart = "2020-03-02", 
           logEnd = "2020-04-02", weight=1)
  plotPred(County = "Polk County", State = "IA", Title = "Polk County, Iowa", 
           logStart = "2020-03-02", logEnd = "2020-04-09", weight = 0)
  plotPred(County = "Oakland County", State = "MI", Title = "Oakland County, Michigan", 
           logStart = "2020-03-02", logEnd = "2020-03-27", weight = 1)
  plotPred(State = "HI", Title = "Hawaii", logStart = "2020-03-02", logEnd = "2020-03-22")
  plotPred(County = "City of St. Louis", Title = "St. Louis (City)", 
           logStart = "2020-03-02", logEnd = "2020-04-06", weight = 2)
  plotPred(County = "St. Louis County", Title = "St. Louis (County)", 
           logStart = "2020-03-02", logEnd = "2020-03-31")
  plotPred(County = "Baltimore City", Title = "Baltimore (City)", 
           logStart = "2020-03-02", logEnd = "2020-03-25")
  plotPred(County = "Durham County", Title = "Durham County", 
           logStart = "2020-03-17", logEnd = "2020-04-03")
  plotPred(County="Miami-Dade County", Title = "Miami-Dade",
           logStart = "2020-03-02", logEnd="2020-03-28")
  plotPred(State="FL", Title = "Florida",
           logStart = "2020-03-02",logEnd="2020-04-01")
  plotPred(State="SD", Title = "South Dakota",
           logStart = "2020-03-02",logEnd="2020-04-02")
  plotPred(State="MT", Title = "Montana",
           logStart = "2020-03-02",logEnd="2020-04-02", weight = 0)
  plotPred(State="IA", Title = "Iowa",
           logStart = "2020-03-02",logEnd="2020-04-02", weight = 0)
  # States with and without statewide public health restrictions
  NoOrders <- c("ND", "SD", "NE", "WY","UT","OK","IA","AR")
  plotPred(State=NoOrders, Title = "States without statewide orders",
           logStart = "2020-03-02",logEnd="2020-04-02", weight = 0)
  plotPred(State=states$abbreviation[!states$abbreviation %in% NoOrders], 
           Title = "States with statewide orders",
           logStart = "2020-03-02",logEnd="2020-04-02", weight = 0)

  # Last week doubling times by state
  STATES <- data.frame(
    abbr = states$abbreviation,
    state = states$state,
    population = 0,
    slope = 0,
    intercept = 0,
    peak = 0,
    k = 0,
    stringsAsFactors = FALSE
  )
  last <- ncol(Cases_USA)
  X <- 1:5
  first <- last - 4
  
  for (i in 1:nrow(STATES))
  {
    STATES$population[i] <- sum(Population_USA$Population[Population_USA$State == STATES$state[i]], na.rm = TRUE)    
    Y <- log(
      colSums(Cases_USA[
      Cases_USA$State == STATES$abbr[i],
      first:last], na.rm=TRUE)
    )
    fit <- lm(Y ~ X)$coefficients
    STATES$intercept[i] <- fit[1]
    STATES$slope[i] <- fit[2]
    if (STATES$slope[i] < 0.01) STATES$slope[i] <- 0.01 
    fit <- covid_fit(Y, log(STATES$population[i] / (1 +  asymptomatic)))
    # Peak can't get higher than 70% of state population
    STATES$intercept[i] <- fit[1]
    STATES$peak[i] <- fit[2]
    STATES$k[i] <- fit[3]
  }
    
  STATES$dtime <- 0.693/STATES$slope
  STATES$dtime[STATES$dtime > 25] <- 25
  STATES$dtime[STATES$dtime < 2] <- 2
  STATES$PredWeek  <- exp(STATES$intercept + (STATES$peak - STATES$intercept) * (1-exp(-STATES$k *  7))) / STATES$population * 100
  STATES$Pred30Day <- exp(STATES$intercept + (STATES$peak - STATES$intercept) * (1-exp(-STATES$k * 30))) / STATES$population * 100
  
  # USA Rate
  USAPopulation <- sum(STATES$population)
  Y <- log(colSums(Cases_USA[, first:last], na.rm=TRUE))
  fit <- lm(Y ~ X)$coefficients
  intercept <- fit[1]
  slope <- fit[2]
  USA <- 0.693/slope
  USAFit <- covid_fit(Y, log(USAPopulation))
  USAPredWeek   <- exp(USAFit[1] + (USAFit[2] - USAFit[1]) * (1-exp(-USAFit[3] *  7))) / USAPopulation * 100
  USAPred30Day  <- exp(USAFit[1] + (USAFit[2] - USAFit[1]) * (1-exp(-USAFit[3] * 30))) / USAPopulation * 100
  
  # By doubling time
  STATES$state <- factor(STATES$state, levels = STATES$state[order(STATES$dtime)], ordered = TRUE)
  Title <- paste("Five day estimation of doubling time", Sys.Date())
  ggObject <- ggplot(STATES, aes(x=state, y=dtime)) +
    geom_col(fill = "brown", color="black", width=.5) +
    theme(axis.text.x=element_text(angle=60, hjust=1)) +
    labs(
      title = Title,
        y = "Doubling Time",
        x = "State"
      ) +
    annotate("segment",x = 0.5, xend = 51.5, y = USA, yend = USA) +
    annotate("text", label=paste("Overall US:", round(USA,1), "days"), x = 1, y=USA, hjust=0, vjust = -0.5)
  nextSlide(ggObject, "Doubling Time")
  
  ggObject <- plot_usmap(data = STATES, values = "dtime", color = "black") +
    scale_fill_continuous(
      low = "red", high = "white", name = "Doubling Time", label = scales::comma
    ) + theme(legend.position = "right") +
    labs(
      title = Title
    )
  nextSlide(ggObject, "Doubling Time")
  
  # Population Percent in five days
  STATES$state <- factor(STATES$state, levels = STATES$state[order(STATES$PredWeek)], ordered = TRUE)
  Title <- paste("Projected cases for", endDate)
  ggObject <- ggplot(STATES, aes(x=state, y=PredWeek)) +
    geom_col(fill = "brown", color="black", width=.5) +
    theme(axis.text.x=element_text(angle=60, hjust=1)) +
  #  scale_y_continuous(breaks = c(0:5)) +
    labs(
      title = Title,
      y = "Percent population in 7 days",
      x = "State"
    ) +
    annotate("segment",x = 0.5, xend = 51.5, y = USAPredWeek, yend = USAPredWeek) +
    annotate("text", label=paste("Overall US:", round(USAPredWeek,1), "percent"), x = 1, y=USAPredWeek, hjust=0, vjust = -0.5)
  nextSlide(ggObject, "All states in 7 days")
  
  ggObject <- plot_usmap(data = STATES, values = "PredWeek", color = "black") +
    scale_fill_continuous(
      low = "white", high = "red", name = "Percent", label = scales::comma
    ) + theme(legend.position = "right") +
    labs(
      title = Title
    )
  nextSlide(ggObject, "Percent population in 7 days")
  
  # Population Percent in 30 days
  STATES$state <- factor(STATES$state, levels = STATES$state[order(STATES$Pred30Day)], ordered = TRUE)
  Title <- paste("Projection for", endDate)
  ggObject <- ggplot(STATES, aes(x=state, y=Pred30Day)) +
    geom_col(fill = "brown", color="black", width=.5) +
    theme(axis.text.x=element_text(angle=60, hjust=1)) +
    #  scale_y_continuous(breaks = c(0:5)) +
    labs(
      title = Title,
      y = "Percent population in 30 days",
      x = "State"
    ) +
    annotate("segment",x = 0.5, xend = 51.5, y = USAPred30Day, yend = USAPred30Day) +
    annotate("text", label=paste("Overall US:", round(USAPred30Day,1), "percent"), x = 1, y=USAPred30Day, hjust=0, vjust = -0.5)
  nextSlide(ggObject, "Percent population in 30 days")
  
  ggObject <- plot_usmap(data = STATES, values = "Pred30Day", color = "black") +
    scale_fill_continuous(
      low = "white", high = "red", name = "Percent", label = scales::comma
    ) + theme(legend.position = "right") +
    labs(
      title = Title
    )
  nextSlide(ggObject, "Percent population in 30 days")
  
  #USA Map of Cases by County
  temp <- Counties
  CROWS <- match(temp$FIPS, Cases_USA$CountyFIPS)
  temp$Cases <- log(Cases_USA[CROWS,ncol(Cases_USA)])
  temp$Cases[is.na(temp$Cases)] <- 0
  Title <- paste("Distribution of Reported Cases as of", today)
  ggObject <- plot_usmap(regions = "counties") +
    geom_point(data = temp, aes(x = LONGITUDE.1, y=LATITUDE.1, size = Cases, color = Cases, alpha = Cases)) +
    scale_color_gradient(low="white",high="red") +
    scale_size(range = c(0, 4)) + 
    scale_alpha(range = c(0, 1)) + 
    theme(legend.position = "none") +
   labs(
      title = Title
    )
  nextSlide(ggObject, "Current US Cases by County")
  
  #USA Map of Deaths by County
  temp <- Counties
  CROWS <- match(temp$FIPS, Deaths_USA$CountyFIPS)
  temp$Deaths <- log(Deaths_USA[CROWS,ncol(Deaths_USA)])
  temp$Deaths[is.na(temp$Deaths)] <- 0
  Title <- paste("Distribution of Deaths as of", today)
  ggObject <- plot_usmap(regions = "counties") +
    geom_point(data = temp, aes(x = LONGITUDE.1, y=LATITUDE.1, size = Deaths, color = Deaths, alpha = Deaths)) +
    scale_color_gradient(low="white",high="red") +
    scale_size(range = c(0, 4)) + 
    scale_alpha(range = c(0, 1)) + 
    theme(legend.position = "none") +
    labs(
      title = Title
    )
  nextSlide(ggObject, "Current US Deaths by County")
  
  
  
  
  # Worldwide
  plotPred(Country = "Italy", Title = "Italy", logStart = "2020-02-22", logEnd = "2020-03-13", weight = 0)
  plotPred(Country = "Spain", Title = "Spain", logStart = "2020-02-25", 
           logEnd = "2020-03-21")
  plotPred(Country = "France", Title = "France", logStart = "2020-02-27", 
           logEnd = "2020-04-04", weight = 0)
  plotPred(Country = "Portugal", Title = "Portugal", 
           logStart = "2020-03-03", logEnd = "2020-03-31")
  plotPred(Country = "Sweden", Title = "Sweden", logStart = "2020-02-28", 
           logEnd = "2020-03-15", weight=0)
  plotPred(Country = "Netherlands", Title = "Netherlands", logStart = "2020-02-29", 
           logEnd = "2020-03-27")
  plotPred(Country = "England", Title = "United Kingdom", 
           logStart = "2020-02-26", logEnd = "2020-03-31", weight = 2)
  plotPred(Country = "South Africa", Title = "South Africa", logStart = "2020-03-08", logEnd = "2020-03-27", weight = 0)
  plotPred(Country = "Brazil", Title = "Brazil", logStart = "2020-03-08", 
           logEnd = "2020-03-31", weight = 1)
  plotPred(Country = "Paraguay", Title = "Paraguay", logStart = "2020-03-09", logEnd = "2020-03-22", weight = 0)
  plotPred(Country = "Rwanda", Title = "Rwanda", 
           logStart = "2020-03-12", logEnd = "2020-03-26")
  plotPred(Country = "Canada", Title = "Canada", logStart = "2020-03-10", logEnd = "2020-03-22")
  plotPred(Country = "Australia", Title = "Australia", logStart = "2020-03-10", 
           logEnd = "2020-03-20", weight=0)
  plotPred(Country = "Germany", Title = "Germany", logStart = "2020-02-26", 
           logEnd = "2020-03-26", weight=0)
  plotPred(Country = "Switzerland", Title = "Switzerland", logStart = "2020-03-10", 
           logEnd = "2020-03-23", weight=1)
  plotPred(Country = "Israel", Title = "Israel", logStart = "2020-02-27", logEnd = "2020-03-25", weight=0)
  plotPred(Country = "Russia", Title = "Russia", logStart = "2020-03-11", logEnd = "2020-04-02", weight=1)
  plotPred(Country = "India", Title = "India", logStart = "2020-03-06", logEnd = "2020-03-25", weight=0)
  plotPred(Country = "Japan", Title = "Japan", logStart = "2020-02-20", 
           logEnd = "2020-04-02", weight=1)
  plotPred(Country = "Mexico", Title = "Mexico", logStart = "2020-02-20", 
           logEnd = "2020-03-30", weight=0)
  
  # World Map of Cases
  CROWS <- match(worldmap$geounit, Cases_Global$Country)
  worldmap$cases <- Cases_Global[CROWS, ncol(Cases_Global)]
  worldmap$lcases <- log(worldmap$cases)
  CROWS <- match(worldmap$geounit, Deaths_Global$Country)
  worldmap$deaths <- Deaths_Global[CROWS, ncol(Deaths_Global)]
  worldmap$ldeaths <- log(worldmap$deaths)
  
  worldmap <- worldmap[!is.na(worldmap$lcases),]
    
  ggObject <-   ggplot() + 
      geom_sf(data = worldmap, color="#00013E", aes(fill=lcases)) +
      scale_fill_gradient(low="#FFCFCF",high="#C00000") +
    labs(
      title = paste("Total Cases as of", Sys.Date())
    ) +
      theme(
        panel.background = element_rect(
          fill = "#00013E"
        ),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        legend.position = "none",
        axis.title.x=element_blank(),
        axis.text.x=element_blank(),
        axis.ticks.x=element_blank(),
        axis.title.y=element_blank(),
        axis.text.y=element_blank(),
        axis.ticks.y=element_blank()
      )
  nextSlide(ggObject, "Worldwide Distribution of Cases")
  
  ggObject <-   ggplot() + 
    geom_sf(data = worldmap, color="#00013E", aes(fill=ldeaths)) +
    scale_fill_gradient(low="#FFCFCF",high="#C00000") +
    labs(
      title = paste("Total Deaths as of", Sys.Date())
    ) +
    theme(
      panel.background = element_rect(
        fill = "#00013E"
      ),
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
      legend.position = "none",
      axis.title.x=element_blank(),
      axis.text.x=element_blank(),
      axis.ticks.x=element_blank(),
      axis.title.y=element_blank(),
      axis.text.y=element_blank(),
      axis.ticks.y=element_blank()
    )
  nextSlide(ggObject, "Worldwide Distribution of Deaths")
  
# Doubling Time by Country
  doubling_Global <- data.frame(
  Country = Cases_Global$Country,
  doublingTime = NA
  )
  
  cols <- (ncol(Cases_Global)-4):ncol(Cases_Global)
  X <- 1:5
  for (i in 1:nrow(doubling_Global))
  {
    if (Cases_Global[i,cols[1]] > 0)
    {
      Y <- unlist(log(Cases_Global[i,cols]))
      doubling_Global$doublingTime[i] <- 0.693 / lm(Y ~ X)$coefficients[2]
    }
  }
  doubling_Global <- doubling_Global[!is.na(doubling_Global$doublingTime),]
  doubling_Global <- doubling_Global[!is.infinite(doubling_Global$doublingTime),]
  doubling_Global$doublingTime <- pmax(doubling_Global$doublingTime, 0)
  doubling_Global$doublingTime <- pmin(doubling_Global$doublingTime, 68)
  doubling_Global$doublingTimeWorldMap <- pmin(doubling_Global$doublingTime, 30)

  CROWS <- match(worldmap$geounit, doubling_Global$Country)
  worldmap$doublingTimeWorldMap <- doubling_Global$doublingTimeWorldMap[CROWS]

  ggObject <-   ggplot() + 
    geom_sf(data = worldmap, color="#00013E", aes(fill=doublingTimeWorldMap)) +
    scale_fill_gradient(high="#FFCFCF",low="#C00000") +
    labs(
      title = paste("Doubling Time as of", Sys.Date())
    ) +
    theme(
      panel.background = element_rect(
        fill = "#00013E"
      ),
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
 #     legend.position = "none",
      axis.title.x=element_blank(),
      axis.text.x=element_blank(),
      axis.ticks.x=element_blank(),
      axis.title.y=element_blank(),
      axis.text.y=element_blank(),
      axis.ticks.y=element_blank()
    )
  nextSlide(ggObject, "Worldwide Doubling Time")
  
  doubling_Global$Country <- factor(doubling_Global$Country, levels = doubling_Global$Country[order(doubling_Global$doublingTime)], ordered = TRUE)
  Title <- paste("Doubling time over the last 5 days as of", Sys.Date())
  
  ggObject <- ggplot(doubling_Global[doubling_Global$doublingTime <= 14,], aes(x=Country, y=doublingTime)) +
    geom_col(fill = "brown", color="black", width=.5) +
    coord_cartesian(
      ylim = c(0,14)) +
    scale_y_continuous(breaks = c(2,4,5,8,10,12,14)) +
    theme(axis.text.x=element_text(angle=90, hjust=1, size = 8)) +
    labs(
      title = Title,
      y = "Doubling Time",
      x = "Country"
    )
  nextSlide(ggObject, "Doubling time (last 5 days)")
  
  ggObject <- ggplot(doubling_Global[doubling_Global$doublingTime > 14,], aes(x=Country, y=doublingTime)) +
    geom_col(fill = "brown", color="black", width=.5) +
    coord_cartesian(
      ylim = c(14,68)) +
    scale_y_continuous(breaks = c(14, 28, 42, 54, 68)) +
    theme(axis.text.x=element_text(angle=90, hjust=1, size = 8)) +
    labs(
      title = Title,
      y = "Doubling Time",
      x = "Country"
    )
  nextSlide(ggObject, "Doubling Time (last 5 days)")

    print(pptx, target = pptxfileName)
} else {
    cat("usafacts.org not updated yet\n")
}
