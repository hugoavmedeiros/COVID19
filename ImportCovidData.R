# Load Data
# USA
SOURCE <- "https://usafactsstatic.blob.core.windows.net/public/data/covid-19/covid_confirmed_usafacts.csv"
Cases_USA <- read.csv(SOURCE, stringsAsFactors = FALSE)
SOURCE <- "https://usafactsstatic.blob.core.windows.net/public/data/covid-19/covid_deaths_usafacts.csv"
Deaths_USA <- read.csv(SOURCE, stringsAsFactors = FALSE)[1:nrow(Cases_USA),]
SOURCE <- "https://www2.census.gov/programs-surveys/popest/datasets/2010-2019/counties/totals/co-est2019-alldata.csv"
Population_USA <- read.csv(SOURCE, stringsAsFactors = FALSE)[, c("STATE", "COUNTY", "STNAME", "CTYNAME", "POPESTIMATE2019")]
names(Population_USA) <- c("STATE", "COUNTY", "State", "County", "Population")

# Notes
# Create states database (FIPS, Abbreviation, Full)
  states <- data.frame(
    FIPS = unique(Cases_USA$stateFIPS)
  )
  states$abbreviation <- Cases_USA$State[match(states$FIPS, Cases_USA$stateFIPS)]
  states$state <- Population_USA$State[match(states$FIPS, Population_USA$STATE)]

# Countys with a Code of 0 match the entire state. Will set population to 0, to match unallocated cases in 
# Cases_USE and Deaths_USA
Population_USA$Population[Population_USA$COUNTY == 0] <- 0
  
# The county FIPS in the usafacts data is the state * 1000 + the county. Need to add this to the
# county population data
Population_USA$CountyFIPS <- Population_USA$STATE * 1000 + Population_USA$COUNTY
Population_USA <- Population_USA[,c("CountyFIPS","County", "State", "Population")]

# Deaths_USA has a repeatd last line. That is why it is 1 record longer. This deletes the extra row
Deaths_USA <- Deaths_USA[1:nrow(Cases_USA),]

# Cases_USA and Deaths_USA are in exactly the same order after above deletion

# The name of CountyFIPS in both files is bizarre, changing it to "CountyFIPS"
names(Cases_USA)[1] <- names(Deaths_USA)[1] <- "CountyFIPS"

# The Wade Hampton Census Area in Cases_USA and Deaths_USA (FIPS 2270) does not show up in the census. This deletes it.
USE <- Cases_USA$CountyFIPS != 2270
Cases_USA <- Cases_USA[USE,]
Deaths_USA <- Deaths_USA[USE,]

# The county called "New York City Unallocated in Cases_USA and Deaths_USA accounts for very few cases.
# I merged it into "New York County"
NYC  <- which(Cases_USA$County.Name == "New York County")
NYCU <- which(Cases_USA$County.Name == "New York City Unallocated")
if (length(NYCU) > 0)
{
  Cases_USA[NYC, 5:ncol(Cases_USA)] <-
  Cases_USA[NYC , 5:ncol(Cases_USA)] +
  Cases_USA[NYCU, 5:ncol(Cases_USA)]
  
  Deaths_USA[NYC, 5:ncol(Deaths_USA)] <-
  Deaths_USA[NYC , 5:ncol(Deaths_USA)] +
  Deaths_USA[NYCU, 5:ncol(Deaths_USA)]
  
  Cases_USA <- Cases_USA[-NYCU,]
  Deaths_USA <- Deaths_USA[-NYCU,]
}
# The county named "Statewide Unallocated" has a countyFIPS of 0. Need to change that to 1000 * stateFIPS
USE <- Cases_USA$CountyFIPS == 0
Cases_USA$CountyFIPS[USE] <- Cases_USA$stateFIPS[USE] * 1000
Deaths_USA$CountyFIPS[USE] <- Deaths_USA$stateFIPS[USE] * 1000

# Verify matching
Cases_USA <- Cases_USA[order(Cases_USA$CountyFIPS),]
Deaths_USA <- Deaths_USA[order(Deaths_USA$CountyFIPS),]
Population_USA <- Population_USA[order(Population_USA$CountyFIPS),]

if (sum(Cases_USA$CountyFips != Deaths_USA$CountyFIPS) > 0)
    cat("Matching failed between allCases and allDeaths\n")
if (sum(Cases_USA$CountyFips != Population_USA$CountyFIPS) > 0)
  cat("Matching failed between allCases and Population\n")

# Global
SOURCE <- "https://github.com/CSSEGISandData/COVID-19/raw/master/csse_covid_19_data/csse_covid_19_time_series/time_series_covid19_confirmed_global.csv"
Cases_Global <- read.csv(SOURCE, stringsAsFactors = FALSE)
SOURCE <- "https://github.com/CSSEGISandData/COVID-19/raw/master/csse_covid_19_data/csse_covid_19_time_series/time_series_covid19_deaths_global.csv"
Deaths_Global <- read.csv(SOURCE, stringsAsFactors = FALSE)

X <- 5:ncol(Cases_Global)
Dates <- gsub("X", "", names(Cases_Global)[X])
DATA <- data.frame(
  Date = Dates,
  Shanghai_Cases = unlist(Cases_Global[Cases_Global == "Shanghai", X]),
  Shanghai_Deaths = unlist(Deaths_Global[Deaths_Global == "Shanghai", X]),
  Beijing_Cases = unlist(Cases_Global[Cases_Global == "Beijing", X]),
  Beijing_Deaths = unlist(Deaths_Global[Deaths_Global == "Beijing", X])
)
write.csv(DATA, file = "Beijing.Shanghai.COVID.csv", row.names = FALSE)
# Notes
# Verified that data are in exactly the same order

# Need to align country names with world map names
worldmap <- ne_countries(
  scale='medium', 
  type = 'map_units', 
  returnclass = 'sf'
)

# Fix country names
Current_Names <- c( 
  "Bahamas",
  "Cabo Verde",
  "Congo (Brazzaville)",
  "Congo (Kinshasa)",
  "Cote d'Ivoire",
  "Diamond Princess",
  "Czechia",
  "Eswatini",
  "Holy See",
  "Korea, South",
  "North Macedonia",
  "Serbia",
  "Taiwan*",
  "US",
  "Timor-Leste",
  "West Bank and Gaza",
  "Guinea-Bissau",
  "Burma",
  "MS Zaandam",
  "United Kingdom",
  "Republic of Serbia",
  "Belgium",
  "Antigua and Barbuda"
)


Revised_Names <- c( 
  "The Bahamas",
  "Cape Verde",
  "Republic of Congo",
  "Democratic Republic of the Congo",
  "Ivory Coast",
  "ship",
  "Czech Republic",
  "Swaziland",
  "Vatican",
  "South Korea",
  "Macedonia",
  "Republic of Serbia",
  "Taiwan",
  "United States of America",
  "East Timor",
  "West Bank",
  "Guinea Bissau",
  "Myanmar",
  "ship",
  "England",
  "Serbia",
  "Flemish Region",
  "Antigua"
)

for (i in 1:length(Current_Names))
{
  Cases_Global$Country.Region[Cases_Global$Country.Region == Current_Names[i]] <- Revised_Names[i]
  Deaths_Global$Country.Region[Deaths_Global$Country.Region == Current_Names[i]] <- Revised_Names[i]
}
# Remove ship
USE <- Cases_Global$Country.Region != "ship"
Cases_Global <-   Cases_Global[USE,]
Deaths_Global <-   Deaths_Global[USE,]
# Verify that Cases_Global$Country.Region are accounted for in worldmap
if (sum(!Cases_Global$Country.Region %in% worldmap$geounit) > 0)
{
  cat("One of the countries in Cases_Global cannot be found in worldmap$geounit\n")
  Cases_Global$Country.Region[!Cases_Global$Country.Region %in% worldmap$geounit]
}

missing <- which(!worldmap$geounit %in% Cases_Global$Country.Region)
# See if missing can be found in the Province.State field. If so, then plug in the country
for (i in 1:length(missing))
{
  X <- which(Cases_Global$Province.State == worldmap$geounit[missing[i]])
  if (length(X) > 0)
  {
#    cat("Setting ",worldmap$geounit[missing[i]], "to",Cases_Global$Country.Region[X],"\n")
    worldmap$geounit[missing[i]] <- Cases_Global$Country.Region[X]
  }
}
# Obvious fixes
worldmap$geounit[worldmap$geounit == "Scotland"] <- "England"
worldmap$geounit[worldmap$geounit == "Wales"] <- "England"
worldmap$geounit[worldmap$geounit == "Jersey"] <- "England"
worldmap$geounit[grep("British", worldmap$geounit)] <- "England"
worldmap$geounit[grep("S.A.R", worldmap$geounit)] <- "China"
worldmap$geounit[grep("Netherlands", worldmap$geounit)] <- "Netherlands"
worldmap$geounit[grep("Brussels", worldmap$geounit)] <- "Flemish Region"
worldmap$geounit[grep("Walloon Region", worldmap$geounit)] <- "Flemish Region"
worldmap$geounit[worldmap$geounit == "Barbuda"] <- "Antigua"
worldmap$geounit[worldmap$geounit == "Puerto Rico"] <- "United States of America"
worldmap$geounit[worldmap$geounit == "Guam"] <- "United States of America"
worldmap$geounit[grep("America", worldmap$geounit)] <- "United States of America"
worldmap$geounit[grep("United States", worldmap$geounit)] <- "United States of America"
worldmap$geounit[grep("French", worldmap$geounit)] <- "France"
worldmap$geounit[worldmap$geounit == "Gaza"] <- "West Bank"
worldmap$geounit[worldmap$geounit == "Aland"] <- "Finland"
worldmap$geounit[worldmap$geounit == "Republic Srpska"] <- "Bosnia and Herzegovina"
worldmap$geounit[worldmap$geounit == "Zanzibar"] <- "Tanzania"
worldmap$geounit[worldmap$geounit == "Northern Island"] <- "England"


# Remove worldmap references not found in Cases_Global
missing <- !worldmap$geounit %in% Cases_Global$Country.Region
worldmap <- worldmap[!missing,]

# Condense cases
Countries <- sort(unique(Cases_Global$Country.Region))
Population <- Deaths <- Cases <- data.frame(
  Country = Countries
)
names <- names(Cases_Global)[5:ncol(Cases_Global)]
Deaths <- Cases <- cbind(Cases, matrix(ncol = length(names), nrow = length(Countries)))
names(Deaths) <- names(Cases) <- c("Country", names)
Population$Population <- 0
for (i in 1:nrow(Cases))
{
  Cases[i,2:ncol(Cases)] <-
    colSums(
      Cases_Global[
        Cases_Global$Country.Region == Cases$Country[i],
        c(5:ncol(Cases_Global))],
      na.rm=TRUE
      )
  Deaths[i,2:ncol(Deaths)] <-
    colSums(
      Deaths_Global[
        Deaths_Global$Country.Region == Deaths$Country[i],
        c(5:ncol(Deaths_Global))],
      na.rm=TRUE
    )
  Population$Population[i] <-
    sum(worldmap$pop_est[
        worldmap$geounit == Cases$Country[i]],
      na.rm=TRUE
    )
}

Cases_Global <- Cases[,1:(allDays+1)]
Deaths_Global <- Deaths[,1:(allDays+1)]
Population_Global <- Population

SOURCE <- "https://www2.census.gov/geo/docs/reference/cenpop2010/county/CenPop2010_Mean_CO.txt"
Counties <- read.csv(SOURCE, stringsAsFactors = FALSE)
Counties$FIPS <- Counties$STATEFP * 1000 + Counties$COUNTYFP
Counties <- Counties[,c("LONGITUDE", "LATITUDE", "FIPS")]
Counties <- usmap_transform(Counties)

# IHME Data
SOURCE <- "https://ihmecovid19storage.blob.core.windows.net/latest/ihme-covid19.zip"
temp <- "IHME.zip"
download.file(SOURCE, temp)
fileNames <- unzip(temp, list=FALSE)
IHME <- read.csv(fileNames[1], stringsAsFactors=FALSE)
