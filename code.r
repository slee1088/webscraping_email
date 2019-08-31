
list.of.packages <- c("rvest","lubridate","dplyr","gmailr","RDCOMClient")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages) 

## Run this if RDCOMClient package cannot be downloaded
# install.packages("RDCOMClient", repos = "http://www.omegahat.net/R") 

library(rvest)
library(lubridate)
library(dplyr)
library(gmailr)

library(RDCOMClient)

if(hour(Sys.time()) < 12){
  greeting <- "Good morning"
} else if (hour(Sys.time()) < 17){
  greeting <- "Good afternoon"
} else {
  greeting <- "Good evening"
}

exch_rate <- read_html("https://www.bloomberg.com/quote/AUDHKD:CUR")

hkd_au <- exch_rate %>%
  html_nodes(".priceText__1853e8a5") %>%
  html_text()

hkd_au_mvt <- exch_rate %>%
  html_nodes(".changeAbsolute__395487f7") %>%
  html_text()

hkd_au_mvt_percent <- exch_rate %>%
  html_nodes(".changePercent__2d7dc0d2") %>%
  html_text()


efinancial_news <- read_html("https://news.efinancialcareers.com/hk-en/en/news-analysis")

top_news <- efinancial_news %>%
  html_nodes(".mt-0 a") %>%
  html_text()

top_news_links <- efinancial_news %>%
  html_nodes(".mt-0 a") %>%
  html_attr("href")

top_news_blurb <- efinancial_news %>%
  html_nodes(".article-byline+ p") %>%
  html_text()

collated <- as.data.frame(cbind(top_news,top_news_links,top_news_blurb)) %>% 
  mutate(mashed=paste0(top_news," - ",top_news_links," (Blurb: ",top_news_blurb,")"))

collated_flattened <- paste(collated$mashed,collapse="\n\n")

subject_body <- paste0(greeting," Scott,\n\n",
                       "The latest AUD-HKD exchange rate is ",hkd_au,".\n",
                       "This is a ",hkd_au_mvt," (",hkd_au_mvt_percent,") change from yesterday.\n\n",
                       "Latest financial news from eFinancial:\n\n",collated_flattened)


subject <- paste0("Daily Update Email - ",Sys.time())

##Sending email via Outlook
# OutApp <- COMCreate("Outlook.Application")
# outMail = OutApp$CreateItem(0)
# outMail[["To"]] = "jung.lee1@cba.com.au"
# outMail[["subject"]] = subject
# outMail[["body"]] = subject_body
# outMail$Send()

##Sending email via Gmail -- refer to link below to setup r so that it can talk to gmail
##https://www.infoworld.com/article/3398701/how-to-send-email-from-r-and-gmail.html

use_secret_file("C:/Users/scott/Downloads/r.json")

test_email <- mime() %>%
  to("slee1088@gmail.com") %>%
  from("slee1088@gmail.com") %>%
  subject(subject) 
test_email$body <- subject_body
send_message(test_email)


