##### UKU SOTG awards analysis

#first install and add these packages (you only need to install once using this code:

install.packages("lme4", "Rmisc", "emmeans", "pbkrtest", "lmerTest", "dplyer")

#then add the packages to your library using this code:

library(lme4)
library(Rmisc)
library(emmeans)
library(pbkrtest)
library(lmerTest)
library(dplyr)

# you can also add them by selecting the check boxes in the 'packages' tab

#next read in the data - lets call it 'df' short for data frame

df <- read.csv("~/uni_2026.csv")

#then we can look at it using the View function
View(df)
               
#we need to make sure R is reading the scores as numbers first so run this code
df$Opposition.Rules.Knowledge.and.Use <- as.numeric(df$Opposition.Rules.Knowledge.and.Use)
df$Opposition.Fouls.and.Body.Contact<- as.numeric(df$Opposition.Fouls.and.Body.Contact)
df$Opposition.Fair.Mindedness <- as.numeric(df$Opposition.Fair.Mindedness)
df$Opposition.Positive.Attitude.and.Self.Control <- as.numeric(df$Opposition.Positive.Attitude.and.Self.Control)
df$Opposition.Communication <- as.numeric(df$Opposition.Communication)


#then we need to get the total scores for each row using this code:
df$sum <- df$Opposition.Rules.Knowledge.and.Use + df$Opposition.Fouls.and.Body.Contact + df$Opposition.Fair.Mindedness + 
  df$Opposition.Positive.Attitude.and.Self.Control + df$Opposition.Communication

# we can plot a histogram of the scores to check they are normally distributed
#and there are no weird outliers using this code:
hist(df$sum) #then switch to the 'plots' tab

#if it all looks good we can calculate the average scores for the clubs

#first lets do this the easy way, just calculate the means and 95% confidence intervals using this code:
#we'll do this for 'club' and we'll call it club means

Unimeans <- group.CI(sum ~ clean_opp, data=df, ci = 0.95)
#lets take a look at the scores
View(Unimeans)
#now lets export these into a csv file so we can read it into excel

write.csv(Unimeans, "~/unimeans.csv")
#this file should appear in the same place as the raw data we imported

#Now we can do this in a more complicated way, which takes into account the structure of the data 
#i.e. who did the scoring at teams scored other teams more than once
#we do this by running a model and extracting the coefficients and predicted means

uni.lm <- lmer(sum ~ clean_opp + (1|clean_team/BreakdownSheetName), data = df)

unimod <- summary(uni.lm)
#this gives coefficients for each team for the model (the first one is always the first alphabetically)
#we just want the model coefficients
#which we can save as a dataframe and export as a csv


unimodcoef <- unimod$coefficients

write.csv(unimodcoef, "~/unimod.csv")
#note that in this csv the top row (named intercept) represents the intercept which all other values are compared to
#the intercept is taken as the club name which is first alphabetically

#finally to check these coefficients line up with the means we can extract the model means 

emm <- emmeans(uni.lm, "clean_opp")

#now we can import that into a csv file and compare the scores and ranks for each
write.csv(emm, "~/unimodmeans.csv")

### work out how many events (remove any with <3)

tournaments <- df %>%
  group_by(BreakdownSheetName, clean_opp) %>%
  dplyr::summarise(n = n())

Scores <-
  df %>%
  group_by(BreakdownSheetName, clean_opp) %>%
  dplyr::summarise(n = n())

Tournaments.entered  <-
  Scores %>%
  group_by(clean_opp) %>%
  dplyr::summarise(n = n())

write.csv(Tournaments.entered, "~/events.csv")

