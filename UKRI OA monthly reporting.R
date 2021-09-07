#OA report for UKRI and Wellcome Trust linked pubs

#Note that source files should be updated before each run of the report.
#Non-compliance report from Symplectic: https://researchpubs.exeter.ac.uk/oareports.html, then select 'Exeter', and select 'compliance report'. No filters.
#Staff List for ref - business objects
#Publications linked to RCUK grants: https://reports.exeter.ac.uk/Reports/Pages/Report.aspx?ItemPath=%2fUoE%2fResearch%2fSymplectic%2fPublications+linked+to+RCUK+Funded+Grants
#New location for above: http://vmetlprod01/ReportS_PAC/report/Symplectic/Reports/Publications%20linked%20to%20RCUK%20Funded%20Grants
#symplectic export - outputs::fetch_sym_report()

#packages------------------
library(dplyr)
library(openxlsx)
library(RColorBrewer)
library(stringr)
library(tidyr)
library(janitor)
library(outputs)
library(ParallelLogger)
#--------------------------

#setup log reporting
logloc <- paste0("//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Standard reports/Report logs/UKRI OA monthly reporting/", Sys.Date(), "_log.txt")

registerLogger(createLogger(name = "DEFAULT_FILE_LOGGER",
                            threshold = "TRACE",
                            appenders = list(createFileAppender(layout = layoutParallel,
                                                                fileName = logloc))))

Start_time <- Sys.time()

#data import (you need to update on each run to capture latest data)------------------------
RCUKlinked <- read.xlsx( 
  "//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Team Information/Source data/Open Access/Publications linked to RCUK Funded grants.xlsx")
  

NonCompliantSixMonths <- read.xlsx( 
  "//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Team Information/Source data/Open Access/Compliance report.xlsx"
  ) %>%
  mutate(`c-embargodate`=as.Date(`c-embargodate`, origin="1899-12-30"),
         `acceptance-date`=as.Date(`acceptance-date`, origin="1899-12-30"),
         `publication-date`=as.Date(`publication-date`, origin="1899-12-30"),
         `embargo-release-date`=as.Date(`embargo-release-date`, origin="1899-12-30")
         )

Staff <- read.xlsx("//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Team Information/Source data/Staff/staff_list.xlsx",
  sheet = "List") %>% 
  rename(Pers.Ref = Per.Ref.No,
         College = Level.2) %>% 
  mutate(Name = paste(Title, Forename, Surname, sep = " "),
         Pers.Ref = as.numeric(Pers.Ref),
         College = str_replace(College, "Office of Vice-Chancellor & Senior Executive", "College of Life & Environmental Sciences"))

Symplectic <- read.xlsx("//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Team Information/Source data/Open Access/Symplectic export.xlsx",
  cols = c(1:120)
  ) %>%
  clean_names() %>% 
  rename(Pers.Ref= users_proprietary_id,
         ORE.handle = open_research_exeter_public_url) %>%
  mutate(Pers.Ref=as.numeric(`Pers.Ref`),
         ID=as.numeric(id),
         doi2 = doi)

DOAJ <-  read.csv("https://doaj.org/csv") 

DOAJ2<-DOAJ %>%
  rename(ISSN=Journal.ISSN..print.version.,
         eISSN=Journal.EISSN..online.version.)


APC1718<-read.xlsx(
  "//universityofexeteruk.sharepoint.com@SSL/DavWWWRoot/sites/OpenResearch/Shared Documents/Open_access/OA_funding_request_spreadsheets/2017-18_OA_funding_requests.xlsx",
  sheet="Jisc APC template v3"
  )

APC1819<-read.xlsx(
  "//universityofexeteruk.sharepoint.com@SSL/DavWWWRoot/sites/OpenResearch/Shared Documents/Open_access/OA_funding_request_spreadsheets/2018-19_OA_funding_requests.xlsx",
  sheet="APC level detail"
  )

APC1920 <- read.xlsx(
  paste0("//universityofexeteruk.sharepoint.com@SSL/DavWWWRoot/sites/OpenResearch/Shared Documents/Open_access/OA_funding_request_spreadsheets/2019-20_OA_funding_requests.xlsx"),
  sheet = "APCs"
  )

APC2021 <- read.xlsx(
  paste0("//universityofexeteruk.sharepoint.com@SSL/DavWWWRoot/sites/OpenResearch/Shared Documents/Open_access/OA_funding_request_spreadsheets/2020-21_OA_funding_requests.xlsx"),
  sheet = "APCs"
  )


APC1718$Date.of.publication <- as.Date(as.numeric(APC1718$Date.of.publication), origin="1899-12-30")
APC1819$Date.of.publication <- as.Date(as.numeric(APC1819$Date.of.publication), origin="1899-12-30")
APC1920$Date.of.publication <- as.Date(as.numeric(APC1920$Date.of.publication), origin = "1899-12-30")
APC2021$Date.of.publication <- as.Date(as.numeric(APC2021$Date.of.publication), origin = "1899-12-30")
colnames(APC1718)[41]<-"blank"

#-----------------------------------------

#join, select, filter------------------------
NonRCUK <- c("British Academy", "British Association of American Studies", "British Geological Survey (BGS)", "Particle Physics and Astronomy Research Council")

CombinedUKRI <- NonCompliantSixMonths %>%
  left_join(select(RCUKlinked,
                   Sponsor,
                   Sponsor_HESA,
                   PUB_ID),
            by=c("Publication.ID" = "PUB_ID")) %>%
  left_join(select(Symplectic,
                   ID,
                   ORE.handle,
                   Pers.Ref,
                   email),
            by=c("Publication.ID" = "ID")
            ) %>%
  filter(Sponsor_HESA=="OST Research Councils" | Sponsor == "Wellcome Trust") %>% 
  filter(!Sponsor %in% NonRCUK) %>% 
  left_join(select(
    APC1718,
    DOI,
    Submitted.by
  ), by=c("doi" = "DOI")
  ) %>%
  mutate(APC1718=case_when(
    is.na(Submitted.by) ~ NA_character_,
    !is.na(Submitted.by) ~ "APC"
  )
  ) %>%
  select(-Submitted.by) %>% 
  left_join(select(
    APC1819,
    DOI,
    Submitted.by
  ), by=c("doi" = "DOI")
  ) %>%
  mutate(APC1819=case_when(
    is.na(Submitted.by) ~ NA_character_,
    !is.na(Submitted.by) ~ "APC"
  )
  ) %>%
  select(-Submitted.by) %>%
  left_join(
    APC1920 %>%
      select(
        DOI,
        Submitted.by
        ),
    by = c("doi" = "DOI")
    ) %>%
  mutate(APC1920 = case_when(
    is.na(Submitted.by) ~ NA_character_,
    !is.na(Submitted.by) ~ "APC"
    )
    ) %>%
  select(-Submitted.by) %>%
  left_join(
    APC2021 %>%
      select(
        DOI,
        Submitted.by
      ),
    by = c("doi" = "DOI")
  ) %>%
  mutate(APC2021 = case_when(
    is.na(Submitted.by) ~ NA_character_,
    !is.na(Submitted.by) ~ "APC"
  )
  ) %>%
  select(-Submitted.by) %>%
  mutate(APC=case_when(
    (is.na(APC1718) & is.na(APC1819) & is.na(APC1920) & is.na(APC2021)) | is.na(doi) ~ NA_character_,
    !is.na(APC1718) | !is.na(APC1819) | !is.na(APC1920) | !is.na(APC2021) ~ "APC"
    )
    ) %>%
  select(-APC1718, -APC1819, -APC1920, -APC2021) %>%
  left_join(select(DOAJ2,
                   ISSN,
                   eISSN,
                   Journal.title),
            by=c("issn" = "ISSN",
                 "eissn" = "eISSN")
  ) %>%
  rename(DOAJ=Journal.title) %>%
  distinct(Publication.ID,
           .keep_all = TRUE) %>% 
  left_join(select(Staff,
                   Pers.Ref,
                   Name,
                   College),
            by="Pers.Ref") %>%
  select(Publication.ID,
         title,
         Sponsor,
         `acceptance-date`,
         `publication-date`,
         `c-embargodate`,
         `embargo-release-date`,
         `c-embargoreason`,
         doi,
         eissn,
         issn,
         journal,
         publisher,
         ORE.handle,
         Item.status.in.repository,
         Name,
         email,
         College,
         APC,
         DOAJ
         ) %>%
  mutate(`Time Since Acceptance`=as.numeric(as.character(Sys.Date()-`acceptance-date`)),
         `Time Since Publication`=as.numeric(as.character(Sys.Date()-`publication-date`))
         ) %>%
  filter(`Time Since Publication`<90) %>%
  arrange(desc(`Time Since Publication`),
          desc(`Time Since Acceptance`),
          Publication.ID,
          email) %>%
  filter(is.na(APC) & is.na(DOAJ)) %>% 
  select(-APC,
         -DOAJ)

#sort out column headings
colnames(CombinedUKRI) <- str_replace_all(colnames(CombinedUKRI),
                                      c("//."=" ",
                                        "-"=" ")
                                      ) 

#-----------------------------------------------

#create a summarised data frame---------------------

SummaryUKRI <- CombinedUKRI %>%
  group_by(Publication.ID,
           title,
           doi,
           journal,
           publisher,
           ORE.handle,
           Sponsor,
           `acceptance date`,
           `publication date`,
           `Time Since Acceptance`,
           `Time Since Publication`) %>%
  summarise(Authors=paste0(unique(Name), collapse="; "),
            Emails=paste0(unique(email), collapse="; "),
            Colleges=paste0(unique(College), collapse="; ")
            ) %>%
  select(Publication.ID,
         title,
         doi,
         journal,
         publisher,
         ORE.handle,
         Sponsor,
         `acceptance date`,
         `publication date`,
         `Time Since Acceptance`,
         `Time Since Publication`,
         Authors,
         Emails,
         Colleges         ) %>%
  arrange(desc(`Time Since Publication`),
          desc(`Time Since Acceptance`)
          )


#------------------------------
  
#set up worksheet/book -------------

#create directory to save exports in
SaveDirectory <- paste0("//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Standard reports/Open Access/UKRI/",Sys.Date())

if (file.exists(SaveDirectory)){
  setwd(file.path(SaveDirectory)) 
  } else  {
    dir.create(file.path(SaveDirectory))
    setwd(file.path(SaveDirectory))
    }


wb <- loadWorkbook(
  "//universityofexeteruk.sharepoint.com/sites/ResPI/analysis/Standard reports/Open Access/UKRI/OA Report template.xlsx"
)


  #create and apply date stamp style
  DateStampStyle<-createStyle(numFmt = "DATE",
                              valign = "top",
                              halign = "left")
  
  #insert date stamp
  writeData(wb,
            "Summary",
            Sys.Date(),
            startCol = 2,
            startRow = 6)
  
  #apply date stamp style
  addStyle(wb,
           "Summary",
           style = DateStampStyle,
           rows = 6,
           cols = 2)
  
  #write data to sheet
  writeDataTable(wb,
                 "Summary",
                 SummaryUKRI,
                 startRow = 10,
                 withFilter = F)
  
  #create and apply text style
  TextStyle <- createStyle(numFmt = "TEXT",
                           wrapText = T,
                           halign = "left",
                           valign = "top")
  
  addStyle(wb,
           "Summary",
           style = TextStyle,
           rows = c(10:(nrow(SummaryUKRI) + 9)),
           cols = c(2:5, 7, 12:14),
           gridExpand = T
           )
  
  #create and apply number style
  NumStyle <- createStyle(numFmt="0",
                          halign="right",
                          valign="top")
  
  addStyle(wb,
           "Summary",
           style=NumStyle,
           rows = c(10:(nrow(SummaryUKRI) + 9)),
           cols = c(1,10:11),
           gridExpand = T
           )
  
  #create and apply date style
  DateStyle<-createStyle(numFmt = "DATE",
                         valign = "top",
                         halign = "right")
  
  addStyle(wb,
           "Summary",
           style = DateStyle,
           rows = c(10:(nrow(SummaryUKRI) + 9)),
           cols = c(8:9),
           gridExpand = T
           )

  
  #create styles (colour settings) for traffic light system
  DayStyle30 <- createStyle(bgFill="#1C74A3",
                            fontColour="#FFFFFF")
  DayStyle60 <- createStyle(bgFill="#8EA683")
  DayStyle90 <- createStyle(bgFill="#FFD762",
                            fontColour="#000000")
  DayStyleLate <- createStyle(fontColour="#FFFFFF",
                              bgFill="#000000")
  DayStyleBlank <- createStyle(bgFill="#CFCFCF")

  
  #add traffic light system - 
  #first since acceptance...
  conditionalFormatting(wb,
                        "Summary",
                        cols=10,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$J11<=30',
                        style=DayStyle30)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=10,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='AND($J11>30,$J11<=60)',
                        style=DayStyle60)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=10,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='AND($J11>60,$J11<=90)',
                        style=DayStyle90)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=10,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$J11>90',
                        style=DayStyleLate)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=10,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$J11=""',
                        style=DayStyleBlank)
  
  #...then for since publication

  conditionalFormatting(wb,
                        "Summary",
                        cols=11,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$K11<=30',
                        style=DayStyle30)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=11,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='AND($K11>30,$K11<=60)',
                        style=DayStyle60)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=11,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='AND($K11>60,$K11<=90)',
                        style=DayStyle90)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=11,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$K11>90',
                        style=DayStyleLate)
  
  conditionalFormatting(wb,
                        "Summary",
                        cols=11,
                        rows=c(11:(nrow(SummaryUKRI) + 10)),
                        rule='$K11=""',
                        style=DayStyleBlank)

  
    #freeze panes
  freezePane(wb,
             "Summary",
             firstActiveRow=11,
             firstActiveCol=1)
  
  #column widths, including hiding UoA details
  setColWidths(wb,
               "Summary",
               cols=c(1),
               widths=13)
  
  setColWidths(wb,
               "Summary",
               cols=c(8:9),
               widths=21)

  setColWidths(wb,
               "Summary",
               cols=c(2:7, 10:14),
               widths="auto")

  #save
  saveWorkbook(wb,
               paste0(SaveDirectory, "/OA non-compliance report.xlsx"),
               overwrite=T)

#close log
logInfo(paste0("Runtime - ", round(Sys.time()-Start_time, 2)))
clearLoggers()
