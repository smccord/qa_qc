library(RODBC)
library(plyr)
library(tidyverse)
library(stringr)
library(reshape2)
library(openxlsx)

data.path <- "C:/Users/samccord/Documents/AIM/Training/2017/LCDO" #set folder path
dima.list <- list.files(path = data.path, pattern = "\\.(MDB)|(mdb)|(accdb)|(ACCDB)$") #create a list of all DIMAs in that path

#Build the DIMA queries that R will call
query.lpi <- "SELECT joinSitePlotLine.SiteName, joinSitePlotLine.SiteID, joinSitePlotLine.PlotID, tblLPIHeader.Observer, joinSitePlotLine.LineID, tblLPIDetail.PointLoc, tblLPIDetail.TopCanopy, tblLPIDetail.Lower1, tblLPIDetail.Lower2, tblLPIDetail.Lower3, tblLPIDetail.Lower4, 
              tblLPIDetail.Lower5, tblLPIDetail.Lower6, tblLPIDetail.Lower7, tblLPIDetail.SoilSurface, 
              tblLPIDetail.HeightWoody, tblLPIDetail.HeightHerbaceous, tblLPIDetail.SpeciesWoody, tblLPIDetail.SpeciesHerbaceous, 
              tblLPIDetail.ChkboxTop,	tblLPIDetail.ChkboxLower1,	tblLPIDetail.ChkboxLower2,	tblLPIDetail.ChkboxLower3,	tblLPIDetail.ChkboxLower4,	tblLPIDetail.ChkboxSoil, tblLPIDetail.ChkboxWoody,	tblLPIDetail.ChkboxHerbaceous,
              tblLPIDetail.ChkboxLower5,	tblLPIDetail.ChkboxLower6,	tblLPIDetail.ChkboxLower7
              FROM joinSitePlotLine INNER JOIN (tblLPIHeader LEFT JOIN tblLPIDetail ON tblLPIHeader.RecKey = tblLPIDetail.RecKey) ON joinSitePlotLine.LineKey = tblLPIHeader.LineKey;"
query.gap <- "SELECT joinSitePlotLine.SiteID, joinSitePlotLine.PlotID, joinSitePlotLine.LineID, tblGapHeader.Observer, tblGapDetail.GapStart, tblGapDetail.GapEnd, tblGapDetail.Gap
FROM joinSitePlotLine INNER JOIN (tblGapHeader INNER JOIN tblGapDetail ON tblGapHeader.RecKey = tblGapDetail.RecKey) ON joinSitePlotLine.LineKey = tblGapHeader.LineKey;"

#Initialize the LPI and Gap data frames
data.lpi <- data.frame()
data.gap <- data.frame()

# For loop to query DIMAs for LPI and Gap data
for (dima.name in dima.list) {
  if (R.Version()[2] == "x86_64") { ## R.Version()[2] is "x86_64" then it will use the function for 64-bit R to create the channel to the DIMA
    dima.channel <- odbcConnectAccess2007(paste(data.path, dima.name, sep = "/")) ## paste() is combining the filepath and DIMA filename, separated with /, to tell odbcConnectAccess2007() where the DIMA is specifically
  } else if (R.Version()[2] == "i386") { ## R.Version()[2] is "i386" then it will use the function for 32-bit R to create the channel to the DIMA
    dima.channel <- odbcConnectAccess(paste(data.path, dima.name, sep = "/")) ## paste() is combining the filepath and DIMA filename, separated with /, to tell odbcConnectAccess() where the DIMA is specifically
  }
  
  data.lpi.current <- sqlQuery(channel = dima.channel,
                               query = query.lpi,
                               stringsAsFactors = F,
                               nullstring = NA)
  data.gap.current <- sqlQuery(channel = dima.channel,
                               query = query.gap,
                               stringsAsFactors = F)
  
  data.lpi <- rbind(data.lpi, data.lpi.current)
  data.gap <- rbind(data.gap, data.gap.current)
  
  odbcClose(channel = dima.channel)
}
data.lpi <- data.lpi %>% filter(!grepl(x = SiteID, pattern = "^demo", ignore.case= T)) #removes all plots under the "Demo" name
data.gap <- data.gap %>% filter(!grepl(x = SiteID, pattern = "^demo", ignore.case= T)) #removes all plots under the "Demo" name

##remove duplicate and random NA fields
data.lpi<-data.lpi[!is.na(data.lpi$SiteID)&!duplicated(data.lpi),]

####LPI Summary####
lpi.summary.observer <- data.lpi %>% group_by(PlotID, Observer) %>%
  summarize(pct.foliar.cover = 100*length(TopCanopy[TopCanopy != "None"])/n(),  #percent total foliar cover
            pct.bare.ground = 100*length(TopCanopy[((TopCanopy == "None") & (SoilSurface == "S") &  #percent bare ground cover
                                                  !grepl(x = Lower1, pattern = "^(L)|(WL)$") &
                                                  !grepl(x = Lower2, pattern = "^(L)|(WL)$") &
                                                  !grepl(x = Lower3, pattern = "^(L)|(WL)$") &
                                                  !grepl(x = Lower4, pattern = "^(L)|(WL)$") &
                                                  !grepl(x = Lower5, pattern = "^(L)|(WL)$"))])/n(),
            pct.litter = 100*length(TopCanopy[(grepl(x = Lower1, pattern = "^(L)|(WL)$") |     # percent litter cover
                                                 grepl(x = Lower2, pattern = "^(L)|(WL)$") |
                                                 grepl(x = Lower3, pattern = "^(L)|(WL)$") |
                                                 grepl(x = Lower4, pattern = "^(L)|(WL)$") |
                                                 grepl(x = Lower5, pattern = "^(L)|(WL)$"))])/n(),
            pct.rock = 100*length(SoilSurface[SoilSurface == "R"| 
                                              SoilSurface == "GR"| 
                                              SoilSurface == "CB"| 
                                              SoilSurface == "ST"| 
                                              SoilSurface == "BY"| 
                                              SoilSurface == "BR"])/n(), #percent rock cover
            pct.embedded.litter = 100*length(SoilSurface[SoilSurface == "EL"])/n(), # percent embedded litter
            pct.moss.lichen.cyano = 100*length(SoilSurface[SoilSurface == "M"|
                                                  SoilSurface == "CY"|
                                                  SoilSurface == "LC"])/n(), # percent moss, cyanobacteria, lichen
            pct.standing.dead=100*length(TopCanopy[ChkboxHerbaceous==T | #percent standing dead hits
                                                     ChkboxLower1==T |
                                                     ChkboxLower2==T |
                                                     ChkboxLower3==T|
                                                     ChkboxLower4==T |
                                                     ChkboxSoil==T |
                                                     ChkboxTop==T|
                                                     ChkboxWoody==T|
                                                     ChkboxLower5==T |
                                                     ChkboxLower6==T |
                                                     ChkboxLower7==T])/n(),
            pct.basal = 100*length(SoilSurface[grepl(x = SoilSurface, pattern = "^[A-Z]{2,4}")])/n()) %>%
  gather(key = "indicator", value = "value", -PlotID, -Observer)

lpi.summary.plot <- lpi.summary.observer %>% ungroup() %>% group_by(PlotID, indicator) %>%
    summarize(value.mean= mean(value), no.lines=n())
  
lpi.summary <- merge(lpi.summary.observer, lpi.summary.plot) %>%
  mutate(value.diff = value - value.mean)

lpi.summary$indicator.title <- lpi.summary$indicator %>% str_replace_all(pattern = "\\.", replacement = " ") %>%
  str_replace(pattern = "^pct", replacement = "Percent ") %>% str_to_title()


####Height Summary####
height.summary.observer<-data.lpi %>% group_by(PlotID, Observer) %>% 
  summarize(avg.woody.height = mean(HeightWoody, na.rm=T),  #average woody height
            avg.herbaceous.height =mean(HeightHerbaceous, na.rm=T),
            n.woody.heights = length(HeightWoody[HeightWoody>=0&!is.na(HeightWoody)]),
            n.herbaceous.heights = length(HeightHerbaceous[HeightHerbaceous>=0&!is.na(HeightHerbaceous)]),
            n.woody0cm = length(HeightWoody[HeightWoody==0&!is.na(HeightWoody)]),
            n.woody1cmto50cm = length(HeightWoody[HeightWoody>=1 &HeightWoody<=50&!is.na(HeightWoody)]),
            n.woody51cmto2m = length(HeightWoody[HeightWoody>50 &HeightWoody<=200&!is.na(HeightWoody)]),
            n.woody2to5m = length(HeightWoody[HeightWoody>200 &HeightWoody<=500&!is.na(HeightWoody)]),
            n.woody5mplus = length(HeightWoody[HeightWoody>500&!is.na(HeightWoody)]),
            n.herbaceous0cm = length(HeightHerbaceous[HeightHerbaceous==0&!is.na(HeightHerbaceous)]),
            n.herbaceous1cmto10cm = length(HeightHerbaceous[HeightHerbaceous>=1 &HeightHerbaceous<=10&!is.na(HeightHerbaceous)]),
            n.herbaceous11cmto30cm = length(HeightHerbaceous[HeightHerbaceous>10 &HeightHerbaceous<=30&!is.na(HeightHerbaceous)]),
            n.herbaceous31cmto50cm = length(HeightHerbaceous[HeightHerbaceous>30 &HeightHerbaceous<=30&!is.na(HeightHerbaceous)]),
            n.herbaceous50cmplus = length(HeightHerbaceous[HeightHerbaceous>50&!is.na(HeightHerbaceous)])) %>%
            
            
  gather(key = "indicator", value = "value", -PlotID, -Observer,-n.woody.heights, -n.herbaceous.heights)

height.summary.observer$value[is.na(height.summary.observer$value)]<-0  #change NAs to 0
height.summary.plot<-height.summary.observer %>% group_by(PlotID, indicator) %>% summarize(no.lines=n(), value.mean=mean(value)) 

height.summary<-merge(height.summary.observer, height.summary.plot)%>%
  mutate(value.diff = value - value.mean)

height.summary$indicator.title <- height.summary$indicator%>% str_replace_all(pattern = "\\.", replacement = " ") %>%
  str_replace(pattern = "^pct", replacement = "Percent of heights") %>% 
  str_replace(pattern = "^n", replacement = "Number of heights")%>%
  str_replace(pattern = "to", replacement = " - ") %>%
  str_replace(pattern = "^avg", replacement = " Average ") 

####Gap Summary####
gap.summary.observer <- data.gap %>% group_by(PlotID, Observer)%>% 
  summarize( n.20to24=length(Gap[(Gap >= 20 & Gap <= 24)]),
              n.25to50 = length(Gap[(Gap >= 25 & Gap <= 50)]),
             n.51to100 = length(Gap[(Gap > 50 & Gap <= 100)]),
             n.101to200 = length(Gap[(Gap > 100 & Gap <= 200)]),
             n.201plus = length(Gap[(Gap > 200)]),
             pct.20to24=100*sum(Gap[(Gap >= 20 & Gap <= 24)])/2500,
             pct.25to50 = 100*sum(Gap[(Gap >= 25 & Gap <= 50)])/2500,
             pct.51to100 = 100*sum(Gap[(Gap > 50 & Gap <= 100)])/2500,
             pct.101to200 = 100*sum(Gap[(Gap > 100 & Gap <= 200)])/2500,
             pct.201plus = 100*sum(Gap[(Gap > 200)])/2500) %>%
  gather(key="indicator", value="value", -PlotID, -Observer)

gap.summary.observer$value[is.na(gap.summary.observer$value)]<-0  #change NAs to 0
gap.summary.plot<-gap.summary.observer %>% group_by(PlotID, indicator) %>% summarize(no.lines=n(), value.mean=mean(value)) 

gap.summary<-merge(gap.summary.observer, gap.summary.plot)%>%
  mutate(value.diff = value - value.mean)

gap.summary$indicator.title <- gap.summary$indicator%>% str_replace_all(pattern = "\\.", replacement = " ") %>%
  str_replace(pattern = "^pct", replacement = "Percent of Line Gaps") %>% str_replace(patter = "^n", replacement = "Number of Gaps")%>%
  str_replace(patter = "to", replacement = " - ") 


####Identify Calibration success
calib.success<-function(df){
  #assign calibration success for percent indicators
  df$success[grepl(x=df$indicator, pattern="^pct") & abs(df$value.diff)>5]<-"Poor"
  df$success[grepl(x=df$indicator, pattern="^pct") & abs(df$value.diff)<=2.5]<-"Good"
  df$success[grepl(x=df$indicator, pattern="^pct")& abs(df$value.diff)<=5& abs(df$value.diff)>2.5]<-"Fair"
  #assign calibration success for average
  df$success[grepl(x=df$indicator, pattern="^avg") & abs(df$value.diff)>4]<-"Poor"
  df$success[grepl(x=df$indicator, pattern="^avg") & abs(df$value.diff)<=2]<-"Good"
  df$success[grepl(x=df$indicator, pattern="^avg")& abs(df$value.diff)<=4& abs(df$value.diff)>2]<-"Fair"
  #assign calibration success for counts in categories
  df$success[grepl(x=df$indicator, pattern="^n") & abs(df$value.diff)>2]<-"Poor"
  df$success[grepl(x=df$indicator, pattern= "^n") & abs(df$value.diff)<=1]<-"Good"
  df$success[grepl(x=df$indicator, pattern="^n")& abs(df$value.diff)<=2 & abs(df$value.diff)>1]<-"Fair"
  
  df
  
} #Assign calibration success function

gap.summary<-calib.success(gap.summary) #gap calibration success
lpi.summary<-calib.success(lpi.summary) #lpi calibration success
height.summary<-calib.success(height.summary)
#####Graph Results####


###Crew Mean Difference Function####
# testing parameters
# df<-lpi.summary
# i<-"n.101to200"
# type<-"n"
crew.mean.difference<-function (df, folder.out){
  a<-0.9  #alpha levels
  
  for (i in unique(df$indicator[df$value.diff!=0])){
    df.indicator<-df[df$indicator==i,] #set up data frame
    
    
    ##identify  percentage of good, fair, poor in calibration
    pct.good<-round(100*nrow(subset(df.indicator, success=="Good"))/nrow(df.indicator))
    pct.fair<-round(100*nrow(subset(df.indicator, success=="Fair"))/nrow(df.indicator))
    pct.poor<-round(100*nrow(subset(df.indicator, success=="Poor"))/nrow(df.indicator))
    
    #set the indicator type
    if (grepl(x=i, pattern="^pct")){
      type<-"pct"
    }else if(grepl(x=i, pattern="^avg")){
      type<-"avg"
    } else  if (grepl(x=i, pattern="^n")){
      type<-"n"
    }else{
       print("There are no calibration standards for this calibration indicator type")
     }
    #set ymin and ymax
    
    if (type=="pct"|type=="avg") {
      r.ymin<-c(-2.5, -5, 2.5)
      r.ymax<-c(2.5, -2.5, 5)
     
      ymax<-round_any(max(abs(df.indicator$value.diff)), 10, f=ceiling)
      ymin<--ymax
      if (ymax<20){
        ymax<-20
        ymin<--20
      }
    }
    
    if (type=="n"){
      r.ymin<-c(-1, -2, 1)
      r.ymax<-c(1, -1, 2)
      #find ymax and ymin for graphs
      ymax<-round_any(max(abs(df.indicator$value.diff)), 2, f=ceiling)
      ymin<--ymax
    }
    #find ymax for graphs             
    ggplot(data = arrange(df.indicator, PlotID, Observer),aes(x = Observer, y = value.diff)) +
      geom_point() +
      scale_y_continuous(limits = c(ymin, ymax)) +
      geom_rect(aes(xmin = 0, xmax = length(unique(Observer)) + 1, ymin = r.ymin[1], ymax = r.ymax[1]),
                # color = "Green",
                fill = "green4", alpha = a) +
      geom_rect(aes(xmin = 0, xmax = length(unique(Observer)) + 1, ymin = r.ymin[2], ymax = r.ymax[2]),
                # color = "Yellow",
                fill = "goldenrod1", alpha = a) +
      geom_rect(aes(xmin = 0, xmax = length(unique(Observer)) + 1, ymin = r.ymin[3], ymax = r.ymax[3]),
                # color = "Orange",
                fill = "goldenrod1", alpha = a) +
      geom_point(data=df.indicator[df.indicator$success=="Poor",],aes(x=df.indicator$Observer[df.indicator$success=="Poor"], 
                                                                            y=df.indicator$value.diff[df.indicator$success=="Poor"]),
                                                                            color="orangered", size=3)+
      geom_point() +
      theme(axis.title.x = element_blank(),
            axis.text.x = element_blank(),
            axis.ticks.x = element_blank(),
            panel.background = element_blank(),
            panel.grid.major.y = element_line(color = "Gray"),
            panel.border = element_rect(color = "Dark gray", fill = NA),
            strip.text.y = element_blank(), 
            plot.caption=element_text(face="bold", size=14))+
      labs(x="Observer", y="Difference from Crew Average", title=unique(df.indicator$indicator.title), 
           caption=paste("Good: ", pct.good, "%   Fair: ", pct.fair, "%   Poor: ", pct.poor, "%" ,sep=""))
     # annotate("text", x=2, y=ymax-2,label= paste("Good: ", pct.good, "%\n Fair: ", pct.fair, "%\nPoor: ", pct.poor, "%" ,sep=""))
    
    ggsave(filename = paste0(folder.out, "/", i,"_mean_diff_figures.png"), device = "png", dpi=600, height=5, width=6)
  }
}
#####
#Crew Mean LPI
crew.mean.difference(df=lpi.summary, folder.out=data.path)
#Crew Mean Gap
crew.mean.difference(df=gap.summary, folder.out=data.path)
#Crew Mean Height
crew.mean.difference(df=height.summary, folder.out=data.path)

###Observer Mean Function####


observer.mean.figures<-function (df, folder.out){
  a<-0.9  #alpha levels
  df$success<-factor(df$success, levels=c("Good", "Fair", "Poor"))
  df$PlotID[df$success=="Good"|df$success=="Fair"]<-paste("*", df$PlotID[df$success=="Good"|df$success=="Fair"], sep="")
  for (i in unique(df$indicator)){
    df.indicator<-subset(df,indicator==i) #set up data frame
    
    ##identify how percentage of good, fair, poor in calibration
    ymin<-0 #find ymin for graphs
    ymax<-round_any(max(df.indicator$value),10, f=ceiling)  #find ymax for graphs             
    ggplot(data = arrange(df.indicator, PlotID, Observer),aes(x = PlotID, y = value)) +
      geom_boxplot(aes(group=PlotID))+geom_point(aes(color=success)) +scale_y_continuous(limits = c(ymin, ymax)) +
      scale_color_manual(values=c( "green4", "goldenrod1","orangered"), name="Difference from \n Crew Mean", drop=F)+
      theme(axis.text.x=element_text(angle=45, hjust=1),
            panel.background = element_blank(),
            panel.grid.major.y = element_line(color = "Gray"),
            panel.border = element_rect(color = "Dark gray", fill = NA),
            strip.text.y = element_blank()) +
      labs(x="Plot", y="Observed Value", title=unique(df.indicator$indicator.title))
      
    ggsave(filename = paste0(folder.out, "/", i,"_observer_mean.png"), device = "png", dpi=600, height=5, width=7)
  }
}
#Observer Mean LPI
observer.mean.figures(lpi.summary, data.path)
#Observer Mean Gap
observer.mean.figures(gap.summary, data.path)
#Observer Mean Height
observer.mean.figures(height.summary, data.path)

####Build a Table of Results####
# Ideas:
#   1) Number of people per crew.mean.difference
#   2) Number of crew.mean.difference
#   3) Mean for each indicator
#   4) Number meeting/not meeting
#   5) Outliers (e.g., gap)
#   6) 

##provides a summary table of the number of groups and observers and the calibration for each method on the whole
calibration.results<-function(lpi=lpi.summary, gap=gap.summary, height=height.summary, file){
  height$method<-"height"
  lpi$method<-"lpi"
  gap$method<-"gap"
  all.summary<-merge(lpi,gap, all=T)
  all.summary<-merge(all.summary, height, all=T)
  all.plots.summary<-all.summary %>% group_by(method) %>% summarise(n.groups=length(unique(PlotID)), n.observers=length(unique(Observer)))
  indicator.total<-all.summary %>% group_by(method) %>% summarise(total=n())
  success<-all.summary %>% group_by(method, success) %>% summarise(n=n()) %>% merge(., indicator.total) %>% mutate(., pct=100*n/total)
  final.summary<-merge(all.plots.summary, success)
  indicator.explanations<- data.frame(indicator=colnames(final.summary), 
                                 comments=c("Calibration method", "Number of groups evaluated for calibration", "Number of observers evaluated for calibration",
                                   "Success category--Good, Fair or Poor", "Number of observer-indicator combinations for each success category", 
                                   "Total number of observer-indicator combinations for that method",
                                   "Percent of observer-indicator combinations for each success category"))
  write.csv(final.summary, file=paste(file, "/calibration.summary.csv", sep=""))
  write.csv(indicator.explanations,file=paste(file, "/calibration.summary.explained.csv", sep="") )
  }

#write out the calibration results
calibration.results(file=data.path)


#calibration results by the individual
all.summary<-merge(lpi.summary,gap.summary, all=T)
all.summary<-merge(all.summary, height.summary, all=T)
write.csv(all.summary, file=paste(data.path,"/calibration.summary.individual.csv", sep=""))
