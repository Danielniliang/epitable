
##################################DISCIPTION################################
#  Functions for data analysis for research projects at Dr.Hackney's Lab   #
#                                                                          #
#  Created by: Liang Ni                                                    #
#  Edited by: Liang Ni                                                     #
#  Last Modified on 1/05/2021                                              #
#                                                                          #
############################################################################

###########################Required packages######################
# install.packages(c("pROC", "car","tableone", "effsize", "dplyr",'readxl',
#                   'readxl','writexl','ggplot2','arsenal','ModelGood'
#                   ,'gmodels','e1071','rio','openxlsx','aod','DescTools'))

library(tableone)
library(car)
library(effsize)
library(dplyr)
library(readxl)
library(writexl)
library(ggplot2)
library(arsenal)
library(ModelGood)
library(gmodels)
library(e1071)
library(rio)
library(stringr)
library(openxlsx)
library(aod)
library(DescTools)

###########################LITTLE BUT USEFUL###################
##Coding missing value into NA##
toNA<-function(vi, list=c('.','MISSING','na',' ','')){
  vi<-as.matrix(vi)
  #set missing value for R:
  for (i in 1:ncol(vi)) {
    col<-vi[,i]
    col[which(col %in% list)]<-NA
    vi[,i]<-col
  }
  vi<-as.data.frame(vi)
  return(vi)
}

# Count the items in a text field separated by certain pattern of string (For example, calculate number of medicine in the field of "Other illness")
count_items <- function(df, var_name, seperate_by = " "){
  var = as.character(df[,c(var_name)])
  count = rep(0, length(var))
  for (i in 1:length(var)) {
    tmp = str_split(var[i], " ")[[1]]
    print(tmp)
    count[i] <- length(tmp[which(nchar(tmp) > 0)])
  }
  return(count)
}

##Calculate N and N missing by group##
N<-function(vi,group){
  #n and nmiss table:
  nmiss<-c()
  for (i in 1:ncol(vi)) {
    tmp<-cbind.data.frame(is.na(vi[,i]),group)
    tmp1<-rbind("",table(tmp$`is.na(vi[, i])`,tmp$group))
    rownames(tmp1)<-c(colnames(vi)[i],"N","Missing")
    nmiss<-rbind(nmiss,"",tmp1)
  }
  return(nmiss)
}

## Change the length of the participant ID into 6 Character ##
To_6Cs_ID<-function(idname, dt){
  participant<-dt[,c(idname)]
  for (i in 1:length(participant)) {
    participant<-as.character(participant)
    id<-participant[i]
    if(nchar(id)==4){id<-paste(substr(id, 1,3), "00", substr(id, 4,nchar(id)), sep = "")}
    if(nchar(id)==5){id<-paste(substr(id, 1,3), "0", substr(id, 4,nchar(id)), sep = "")}
    if(nchar(id)==6){id=id}
    participant[i]<-id
  }
  dt[,c(idname)]<-participant
  return(dt)
}

##Formating and Genrating Excel Table##

##Arguments##
#table        #A result table
#title        #Title of the table. A character.
#caption      #Caption at the bottom of the table. A character.
#outfile      #Output file location
#openfile     #Open the file or not
##Function##
ExcelTable<-function(table, 
                     title="Table.N",
                     caption="haha",
                     outfile=outfile,
                     openfile=T){
  table<-rbind(as.matrix(cbind(rownames(table),table)),
               c(caption,
                 rep("",ncol(table))))
  ## Create a workbook
  x<-rbind(colnames(table),table)
  colnames(x)<-c(title,rep("",ncol(x)-1))
  #write and designate a workbook
  wb <- write.xlsx(x, outfile)
  
  #Customize style:
  First2<-createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin',textDecoration = "bold")
  Border <- createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin')
  bold<-createStyle(textDecoration = "bold")
  wraptext<-createStyle(wrapText = T,valign = 'top',halign = 'left')
  #Apply style:
  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, Border, rows = nrow(x)-1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, Border, rows = nrow(x), cols = 1:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, bold, rows = 3:nrow(x)-3, cols = 1, gridExpand = T)
  
  ## Merge cells: Merge the footer's rows
  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  #mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x))
  mergeCells(wb, 1, cols = 1:ncol(x), rows = (nrow(x)+1):(nrow(x)+4))
  addStyle(wb,1,wraptext,cols = 1:ncol(x), rows = nrow(x)+1)
  
  #set colunme width:
  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto",ignoreMergedCells = T)
  #setRowHeights(wb, 1, rows = nrow(x)+1, heights  = "auto")
  
  #Modify font:
  modifyBaseFont(wb, fontSize = 12, fontColour = "black",fontName = "Times New Roman")
  
  ## Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)
  
  if(openfile==T){
    openXL(outfile)
  }
  return(wb)
}

########################### Normality Test and Distribution Plots #############################
##Desciption##
#Function for generating a table for normality test and generate distribution plots

#Packages Needed:e1071,ggplot2

##Arguments##
#dt              #Matrix or dataframe including variables of interest
#plot            #If generate plots
#table           #If generate tables
#group           #Treatment groups (Exposures). A vector.
#plot.filename   #File name for plots
#Table.filename  #File name for the result table

##Function##
dist.plots<-function(dt,group,
                     ks=F,
                     plot=T,
                     table=T,
                     plot.filename='Distribution Plots for Variables',
                     Table.filename='Normality Test Results'){

  #Shapiro test, Kurtosis and Skewness
  dt<-apply(dt,2,as.numeric)
  if(table==T){
    result1<-c()
    for (i in 1:ncol(dt)) {
      if(sum(is.na(dt[,i])==F)<3){next}
      if(ks==F){
        re<-shapiro.test(as.numeric(dt[,i]))
      }else{
        re<-ks.test(dt[,i], "pnorm", mean=mean(dt[,i],na.rm = T), sd=sd(dt[,i], na.rm=T))
      }
      
      kt<-e1071::kurtosis(as.numeric(dt[,i]),na.rm=TRUE)
      skew<-e1071::skewness(as.numeric(dt[,i]),na.rm=TRUE)
      result1<-rbind(result1, round(c(re$statistic,re$p.value,kt,skew),4))
    }
    #Summary of variables
    gt<-group_by(cbind.data.frame(dt,group) , group)
    sum_result<-c()
    for (i in 1:(ncol(gt)-1)) {
      sum<-summarize_at(gt[,1:(ncol(gt)-1)],names(gt)[i],funs( mean(.,na.rm=T),
                                                               sd(.,na.rm=T),
                                                               median(.,na.rm=T),
                                                               IQR(.,na.rm=T),
                                                               min(.,na.rm=T),
                                                               max(.,na.rm=T)
      ))
      N<-sum(is.na(gt[,i])==F)
      N_missing<-sum(is.na(gt[,i]))
      sum_result<-rbind(sum_result,cbind(N,N_missing,sum))
    }
    result<-cbind(colnames(dt),sum_result,result1)
    colnames(result)<-c('Variables',colnames(sum_result),
                        ifelse(ks==T,
                               'Kolmogorov-Smirnov test statistics',
                               'Shapiro test statistics')
                        ,
                        'P value','Kurtosis','Skewness')
    write.csv(result,paste(Table.filename,'.csv'))
  }


  if(plot==T){
    dt<-as.data.frame(dt)
    #Distribution Ploting
    pdf(paste(plot.filename,".pdf"))
    for (i in 1:ncol(dt)) {
      #Histogram and Density Plot
      p1<-ggplot(dt, aes(x=as.numeric(dt[,i]))) +
        geom_histogram(aes(y=..density..),
                       # Histogram with density instead of count on y-axis
                       color='Black',
                       fill="#56B4E9") +
        geom_density(alpha=.2)+
        labs(title=paste("Histogram and Density Curve of",names(dt)[i]),
             x=names(dt)[i],
             y = "Density")
      #Density plot by groups:
      p2<-ggplot(dt, aes(x=as.numeric(dt[,i]) , fill=group)) +
        geom_density(alpha=.3)+
        labs(title=paste("Density Curve of",names(dt)[i]),
             x=names(dt)[i],
             y = "Density",
             fill=group)

      #QQ-PLot
      p3<-ggplot(dt, aes(sample = as.numeric(dt[,i]))) +
        stat_qq() +
        stat_qq_line()+
        labs(title=paste("Q-Q plot of",names(dt)[i]),
             x= group,
             y = names(dt)[i])
      #QQ-Plot by group
      p4<-ggplot(dt, aes(sample = as.numeric(dt[,i]), colour = factor(group))) +
        stat_qq() +
        stat_qq_line()+
        labs(title=paste("Q-Q plot of",names(dt)[i]),
             x= group,
             y = names(dt)[i])

      #Box plot(with points)
      p5<-ggplot(dt, aes(y=as.numeric(dt[,i]),x=group,color=group) ) +
        geom_boxplot()+
        labs(title=paste("Box plot of",names(dt)[i]),
             x= group,
             y = names(dt)[i])
      print(p1)
      print(p2)
      print(p3)
      print(p4)
      print(p5)
    }
    dev.off()
  }

  return(result)
}


########################### Demographic Characteristic Table ###################################
##Desciption##
#Function for generating demographic characteristic tables or Univariate analysis

#Packages #Packages needed: dplyr, openxlsx

##Arguments##
#con            #Matrix including continues variables
#cat            #Matrix including categorical variables
#group          #Treatment groups (Exposures). A vector.
#median         #Whether to present median results (with Mann-Whitney U test for difference between group)
#Fisher         #Whether to use Fisher's Exact test for categorical variables)
#Ncolumn        #Whether to add N columns for each treatment groups
#outfile        #Output file location

##Function##
DemoTable<-function(con=c(),
                    cat=c(),
                    group,
                    median=F,
                    Fisher=F,
                    Ncolumn=F,
                    paired=F,
                    decimal=2,
                    outfile='Descriptive Analysis Result.xlsx'){
  if(length(con)>0){
    # Tables for continuous variabels
    data<-as.data.frame(apply(apply(con,2,as.character),2,as.numeric))
    da<-as.data.frame(cbind(data,group))
    gt<-group_by(da, group)
    group<-factor(group)
    #Tables for mean ± SD, P by ANOVA
    if(median==F){
      pval<-c()
      e<-c()
      x=0
      for (i in 1:ncol(data)) {
        x=x+1
        sum <- summarize_at(gt,variable.names(data)[i],list(~mean(.,na.rm=T),~sd(.,na.rm=T)))
        mean_sd<-apply(as.matrix(sum)[,2:3], 2, as.numeric)
        #tmp<-paste(round(mean_sd[,1],2),'±', round(mean_sd[,2],1),sep=" ")
        m<-paste(format(round(mean_sd[,1],decimal),nsmall=decimal),'±', format(round(mean_sd[,2],decimal),nsmall = decimal),sep=" ")
        #m1<-c(paste(round(mean(data[,i],na.rm = TRUE),2),'±', round(sd(data[,i],na.rm = TRUE),1),sep=" "),m)
        m1<-c(paste(format(round(mean(data[,i],na.rm = TRUE),decimal), nsmall = decimal),'±', format(round(sd(data[,i],na.rm = TRUE),decimal),nsmall = decimal),sep=" "),m)

        e<-rbind(e,m1)
      #Calculate P value using t test or One way ANOVA
        x=x+1
        sum1 <- summarize_at(gt,variable.names(data)[i],funs(mean(.,na.rm=T),sd(.,na.rm=T)))
        if(is.na(sum(sum1[,2]))){pv<-NA
        }else{
          if(paired==T){
            pv<-round(t.test(data[which(group==levels(group)[1]),i],data[which(group==levels(group)[2]),i] ,paired=T)$p.value,3)
          }else{
            if(nlevels(group)>2){
              temp<-Anova(lm(data[,i]~group))
              # test<-unlist(summary(temp))
              pv<-round(temp[1,4],3)
            }else{
              Ftest<-var.test(data[,i]~group)
              if(Ftest$p.value<0.05){
                test<-t.test(as.numeric(data[,i])~group,var.equal=F)
              }else{test<-t.test(as.numeric(data[,i])~group,var.equal=T)
              }
              pv<-round(test$p.value,3)
            }
          }
        }
        
        if(is.na(pv)){pval<-rbind(pval,"NA")
        }else{
          if(pv<0.05& pv>=0.01){
            pval<-rbind(pval,paste(pv,'*'))
          }else{
            if(pv<0.01 & pv>=0.001){
              pval<-rbind(pval,paste(pv,'**'))
            }else{
              if(pv>=0.05){
                pval<-rbind(pval,paste(pv,""))
              }else{
                if(pv<0.001){
                  pval<-rbind(pval,"<.001 **")
                }else{
                  pval<-rbind(pval,"NA")
                }
              }
            }
          }
        }
      }
      totaln<-sum(table(group))
      np<-table(group)
      n<-c(totaln,np,"")
      e<-rbind(n,cbind(e,pval))
      rownames(e)<-c("n",variable.names(data))
      colnames(e)<-c("Total",levels(factor(group)),"p-value")
    }else{

      #Tables for Median (IQR), P by U test (For study group = 2)
      pval<-c()
      e<-c()
      x=0
      for (i in 1:ncol(data)) {
        x=x+1
        sum <- summarize_at(gt,variable.names(data)[i],list(~median(.,na.rm=T),~IQR(.,na.rm=T)))
        m<-c()
        for (j in 1:length(levels(factor(group)))) {
          tmp<-paste(round(sum[j,2],1),' (', round(sum[j,3],2),')',sep="")
          m<-c(m,tmp)
        }
        m1<-c(paste(round(median(data[,i],na.rm = TRUE),1),' (', round(IQR(data[,i],na.rm = TRUE),2),')',sep=""),m)

        e<-rbind(e,m1)
      #Calculate P value using Mann Whitney U test or Kruskal Wallis Test One Way Anova
        x=x+1
        sum1 <- summarize_at(gt,variable.names(data)[i],funs(mean(.,na.rm=T),sd(.,na.rm=T)))
        
        if(is.na(sum(sum1[,2]))){
          pv<-NA
        }else {
          if(nlevels(group)>2){
            pv<-round(kruskal.test(data[,i]~group)$p.value ,3)
          }else{
            if(paired==T){
              pv<-round(wilcox.test(data[which(group==levels(group)[1]),i],data[which(group==levels(group)[2]),i] ,paired=T)$p.value,3)
            }else{
              pv<-round(wilcox.test(data[,i]~group)$p.value,3)
            }
            
          }
        }
          
        if(is.na(pv)){pval<-rbind(pval,"NA")
        }else{
          if(pv<0.05& pv>=0.01){
            pval<-rbind(pval,paste(pv,'*'))
          }else{
            if(pv<0.01 & pv>=0.001){
              pval<-rbind(pval,paste(pv,'**'))
            }else{
              if(pv>=0.05){
                pval<-rbind(pval,paste(pv,""))
              }else{
                if(pv<0.001){
                  pval<-rbind(pval,"<.001 **")
                }else{
                  pval<-rbind(pval,"NA")
                }
              }
            }
          }
        }
        }

      totaln<-sum(table(group))
      pct<-round(prop.table(table(group))*100,2)
      np<-paste(table(group),' (',pct,')', sep='')
      n<-c(totaln,np,"")

      e<-rbind(n,cbind(e,pval))
      rownames(e)<-c("n",variable.names(data))
      colnames(e)<-c("Total",levels(factor(group)),"p-value")
    }
  }

  if(length(cat)>0){
    #Tables for categorical variabes
    data1<-as.data.frame(apply(cat,2,factor))
    f<-c()
    x=0
    for (i in 1:ncol(data1)) {
      if(length(data1[,i][which(is.na(data1[,i])==T)])==length(data1[,i])){next}
      x=x+1
      m <-t(rbind(table(data1[,i]),table(group,data1[,i])))
      #rownames(m)<-c(variable.names(data1)[i],rep(" ",length(levels(factor(data1[,i])))))
      n<-table(group,data1[,i])
      cols<-c()
      for (j in 1:ncol(m)) {
        cols<-c(cols,sum(m[,j]))
      }
      #Formating and Calculating P values
      for (j in 1:nrow(m)){
        for (k in 1:ncol(m)){
          m[j,k] = paste(m[j,k],' (',round(as.numeric(m[j,k])*100/cols[k],1),')',sep = '')
        }
        if(Fisher==F){
          pv<-chisq.test(n)
        }else{
          pv<-fisher.test(n)
        }
        if(pv$p.value<0.001){
          pval<-"<.001 **"
        }else{
          if(pv$p.value<0.01){
            pval<-paste(round(pv$p.value,3),"**")
          }else{
            if(pv$p.value<0.05& pv$p.value>=0.01){
              pval<-paste(round(pv$p.value,3),"*")
            }else{
              pval<-round(pv$p.value,3)
            }
          }
        }
      }
      m<-cbind(rownames(m),m,"")
      m<-rbind(c(rep("",ncol(m)-1),pval),m)


      rownames(m)<-c(variable.names(data1)[i],rep(" ",nrow(m)-1))
      f<-rbind(f, "", m)
    }
    colnames(f)<-c("","Total",levels(factor(group)),"")
  }

  
  
  #option of adding N columns for each group:
  #Function of adding N columns:
  if(Ncolumn==T){
    Ntable<-function(vi,result,group){
      n<-c()
      for (i in 1:ncol(vi)) {
        tmp<-cbind.data.frame(is.na(vi[,i]),group)
        tmp1<-table(tmp[,1],tmp[,2])
        n<-rbind(n,tmp1[1,])
      }
      rownames(n)<-colnames(vi)
      
      nresult<-c()
      for (i in 1:ncol(n)) {
        temp<-cbind(c("",n[,i]),result[,i+1])
        nresult<-cbind(nresult,temp)
      }
      nresult<-cbind(result[,1],nresult,result[,ncol(result)])
      level<-c()
      for (i in 1:nlevels(factor(group))) {
        level<-cbind(level,cbind('N',levels(factor(group))[i]))
      }
      colnames(nresult)<-c('Total',level,"P value")
      return(nresult)
    }
    e<-Ntable(vi=con,result = e,group = group)
  }
    
  
  if(length(con)>0&length(cat)==0){
    table<-e
    }else{
      if(length(cat)>0&length(con)==0){
        table<-f
      }else{
        if(Ncolumn==T){
          
          ftmp=c()
          for (i in 2:(ncol(f)-2)) {
            ftmp=cbind(ftmp,cbind(f[,i],""))
          }
          f=cbind(f[,1],ftmp,f[,(ncol(f)-1):ncol(f)])
          table<-rbind(cbind("",e),f)
        }else{
          table<-rbind(cbind("",e),f)
        }
      }
    }


  y<-table 
  #Notation
  if(is.null(con)==F & is.null(cat)==F){
    table<-rbind(cbind(rownames(table),table),
                 c(
                   paste("Table 1. Values are presented as",
                         ifelse(median,"Median (IQR)","Mean ± SD"),
                         "for continuous variables, and n (%) for categorical variables. P values were calculated with" ,
                         ifelse(median,
                                ifelse(nlevels(group)!=2, "Kruskal Wallis Test One Way ANOVA","Mann-Whitney U test"),
                                ifelse(nlevels(group)!=2, "One-way ANOVA",ifelse(paired,"Paired t test","independent t test"))),
                         "for continuous variables and",
                         ifelse(Fisher, "Fisher's exact test", "Chi-square test"),
                         "for categorical variables."),
                   rep("",ncol(table))
                 ))
  }else{
    if(is.null(con)==T & is.null(cat)==F){
      table<-rbind(cbind(rownames(table),table),
                   c(
                     paste("Table 1. Values are presented as n (%). P values were calculated with" ,
                           ifelse(Fisher, "Fisher's exact test", "Chi-square test")),
                     rep("",ncol(table))
                   ))
    }else{
      if(is.null(con)==F & is.null(cat)==T){
        table<-rbind(cbind(rownames(table),table),
                     c(
                       paste("Table 1. Values are presented as",
                             ifelse(median,"Median (IQR).","Mean ± SD."),
                             "P values were calculated with" ,
                             ifelse(median,
                                    ifelse(nlevels(group)!=2, "Kruskal Wallis Test One Way ANOVA","Mann-Whitney U test"),
                                    ifelse(nlevels(group)!=2, "One-way ANOVA",ifelse(paired,"Paired t test","independent t test")))),
                       rep("",ncol(table))
                     )) 
      }
    }
  }
  


 # export(as.data.frame(table),outfile)
  ## Create a workbook
  x<-rbind(colnames(table),table)
  colnames(x)<-rep(paste("Table 1. Demographic Characteristics",levels(group)[1],"vs.",levels(group)[2]),ncol(x))
  #write and designate a workbook
  wb <- write.xlsx(x, outfile)

  #Customize style:
  First2<-createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin',textDecoration = "bold")
  Border <- createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin')
  bold<-createStyle(textDecoration = "bold")
  wraptext<-createStyle(wrapText = T,valign = 'top',halign = 'left')
  #Apply style:
  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = T)
  if(Ncolumn==T){
    addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = T)
  }else{
    addStyle(wb, sheet = 1, First2, rows = 2, cols = 3:ncol(x), gridExpand = T)
  }
  #addStyle(wb, sheet = 1, Border, rows = nrow(x)-1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, Border, rows = nrow(x), cols = 1:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, bold, rows = 3:nrow(x)-3, cols = 1, gridExpand = T)

  ## Merge cells: Merge the footer's rows
  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  #mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x))
  mergeCells(wb, 1, cols = 1:ncol(x), rows = (nrow(x)+1):(nrow(x)+4))
  addStyle(wb,1,wraptext,cols = 1:ncol(x), rows = nrow(x)+1)

  #set colunme width:
  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto",ignoreMergedCells = T)
  #setRowHeights(wb, 1, rows = nrow(x)+1, heights  = "auto")

  #Modify font:
  modifyBaseFont(wb, fontSize = 12, fontColour = "black",fontName = "Times New Roman")

  ## Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)

  openXL(outfile)
  return(y)
}


###########################Univaraite Analysis (Test for group differnces)##################
##Desciption##
#Function for testing of difference between group

#Packages #Packages needed: dplyr, openxlsx

##Arguments##
#dt             #Matrix including variables of interest (continues varuavkes)
#group          #Treatment groups (Exposures). A vector.
#median         #Whether to present median results (with Mann-Whitney U test for difference between group)
#outfile        #Output file location

##Function##
group.diff<-function(dt,group,
                     median=F,
                     paired=F,
                     cohend.correct=T,
                     outfile='Comparison of difference between group result.xlsx'){
  #Transform the data into numeric
  dt<-as.data.frame(apply(dt,2,as.numeric))

  #calculate the mean and standard deviation for each group
  da<-group_by(cbind.data.frame(dt,group),group)
  result1<-as.data.frame(summarise_all(da,funs(mean(.,na.rm=T),sd(.,na.rm=T),median(.,na.rm = T),IQR(.,na.rm = T)))[,-1])
  rownames(result1)<-levels(group)
  mean<-result1[,1:ncol(dt)]
  sd<-result1[,(ncol(dt)+1):(2*ncol(dt)),]
  Median<-result1[,(2*ncol(dt)+1):(3*ncol(dt)),]
  iqr<-result1[,(3*ncol(dt)+1):(4*ncol(dt)),]
  mean_dff<-unlist(mean[2,]-mean[1,])

  #Testing:
  #Testing for continues Var

  result2<-c()
  for (i in 1:ncol(dt)) {
    # N missing:
    N_missing<-sum(is.na(dt[,i]))
    if(median==F){

      #t test:
      if(paired==T){
        group<-factor(group)
        test<-t.test(dt[which(group==levels(group)[1]),i],dt[which(group==levels(group)[2]),i] ,paired=T)
      }else{
        Ftest<-var.test(dt[,i]~group)
        if(Ftest$p.value<0.05){
          test<-t.test(as.numeric(dt[,i])~group)
        }else{test<-t.test(as.numeric(dt[,i])~group,var.equal=TRUE)
        }
      }

    }else{
      #U test:
      if(paired==T){
        test<-wilcox.test(dt[which(group==levels(group)[1]),i],dt[which(group==levels(group)[2]),i] ,paired=T)
      }else{
        test<-wilcox.test(as.numeric(dt[,i])~group)
      }
      
    }

    groupi<-relevel(group,ref = levels(group)[2])
    #Cohen's d
    if(cohend.correct==T){
      if(paired==F){
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=FALSE,
                       na.rm=FALSE, hedges.correction=TRUE,
                       conf.level=0.95,noncentral=FALSE)
      }else{
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=TRUE,
                       na.rm=FALSE, hedges.correction=TRUE,
                       conf.level=0.95,noncentral=FALSE)
      }
    }else{
      if(paired==F){
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=FALSE,
                       na.rm=FALSE, hedges.correction=FALSE,
                       conf.level=0.95,noncentral=FALSE)
      }else{
        cohen<-cohen.d(as.numeric(dt[,i])~groupi,pooled=TRUE,paired=TRUE,
                       na.rm=FALSE, hedges.correction=FALSE,
                       conf.level=0.95,noncentral=FALSE)
      }
    }

    #Formatting:
    cd<-paste(round(cohen$estimate,2)," (",round(cohen$conf.int[1],1), " ,",round(cohen$conf.int[2],1),")",sep = "")



    if(median==F){
      meandiff<-paste(round(mean_dff[i],2)," (",round(test$conf.int[1],2), " ,",round(test$conf.int[2],2),")",sep = "")
      result2<-rbind(result2, c(meandiff, round(test$p.value,3), cd, N_missing))
    }else{
      result2<-rbind(result2, c(round(test$p.value,3), cd, N_missing))
    }
  }

  if(median==F){
    result<-cbind(round(t(mean),2),result2)
    result[,1:2]<-cbind(paste(result[,1], "±",round(sd[1,],2)),paste(result[,2], "±",round(sd[2,],2)))
    rownames(result)<-colnames(dt)
    #colnames(result)<-c(paste(levels(group)[1],"(Mean ± SD)"),paste(levels(group)[2],"(Mean ± SD)"),"Mean Difference",
    #                    "CI Lower","CI Upper","statistics","P-value","cohen's D","CI Lower","CI Upper","Missing")
  }else{
    result<-cbind(round(t(Median),2),result2)
    result[,1:2]<-cbind(paste(result[,1], " (",round(iqr[1,],2),")",sep = ""),paste(result[,2], " (",round(iqr[2,],2),")",sep = ""))
    rownames(result)<-colnames(dt)
    #colnames(result)<-c(paste(levels(group)[1],"(Median (IQR))"),paste(levels(group)[2],"(Median (IQR))"),
    #                    "statistics","P-value","cohen's D","CI Lower","CI Upper","Missing")
  }
  #N value for result:
  #N:
  vi<-dt
  n<-c()
  for (i in 1:ncol(vi)) {
    tmp<-cbind.data.frame(is.na(vi[,i]),group)
    tmp1<-table(tmp$`is.na(vi[, i])`,tmp$group )
    tmp2<-sum(is.na(vi[,i])==F)
    n<-rbind(n,c(tmp2,tmp1[1,]))
  }
  result4<-cbind(n[,2], result[,1],n[,3],result[,2],result[,3:ncol(result)],n[,1])

  #Naming columns:
  if(median==F){
    colnames(result4)<-c("N",paste(levels(group)[1],"(Mean ± SD)"),"N",paste(levels(group)[2],"(Mean ± SD)"),"Mean Difference",
                         "P-value","Cohen's D (CI)","N Missing","Total N")
  }else{
    colnames(result4)<-c("N",paste(levels(group)[1],"(Median (IQR))"),"N",paste(levels(group)[2],"(Median (IQR))"),
                         "P-value","Cohen's D (CI)","N Missing","Total N")
  }

  y<-result4
  ####Formating worksheet#####
  #Title:
  result4<-cbind(rownames(result4),result4)
  x<-rbind(colnames(result4),result4)
  colnames(x)<-rep(paste("Table 2. ",levels(group)[1],"vs.",levels(group)[2]),ncol(x))

  #Footnote:
  if(median==F){
    x<-rbind(x,
             "Comparison of outcome variables between. P values were obtained by independent t-test. ")
  }else{
    x<-rbind(x,
             "Comparison of outcome variables between. P values were obtained by Mann-Whitney U test")
  }


  #write and designate a workbook
  wb<-createWorkbook()
  addWorksheet(wb, "Analysis Results")
  writeData(wb,1,x)
  #wb <- write.xlsx(x, outfile)
  #Customize style:
  First2<-createStyle(border="Bottom", borderColour = "black",borderStyle = 'thin',textDecoration = "bold")
  Border <- createStyle(border="TopBottom", borderColour = "black",borderStyle = 'thin')
  bold<-createStyle(textDecoration = "bold")
  #Apply style:
  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, Border, rows = nrow(x)-1, cols = 1:ncol(x), gridExpand = T)
  addStyle(wb, sheet = 1, Border, rows = nrow(x)+1, cols = 1:ncol(x), gridExpand = T)
  #addStyle(wb, sheet = 1, bold, rows = 3:nrow(x)-3, cols = 1, gridExpand = T)

  ## Merge cells: Merge the footer's rows
  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  #mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x))
  mergeCells(wb, 1, cols = 1:ncol(x), rows = nrow(x)+1)

  #set colunme width:
  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto",ignoreMergedCells = T)

  #Modify font:
  modifyBaseFont(wb, fontSize = 12, fontColour = "black",fontName = "Times New Roman")

  ## Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)

  openXL(outfile)

  return(y)
}



###########################Multivariate Analysis Result Tables#################################
##Desciption##
#Function to Generate Mutivariate analaysis result table by use of multivariate linear regression model

#Packages Needed:dplyr,openxlsx

##Arguments##
#dt           #Dataset including all variables
#vi_name      #Outcome variables
#group        #Treatment groups (Exposures). A Vector or a variable names in dt
#con_name     #vector of Variable names
#cat_name     #vector of Variable names
#ref          #Reference group. A level name
#method       #if methd="logistic", logistic regression will be generated with Odds ratios presented
#outfile      #Output file location

##Function##
multivariate<-function(dt,
                       group,
                       vi_name,
                       con_name,
                       cat_name,
                       ref=c(),
                       method="Linear",
                       CI=F,
                       log.transform=F,
                       outfile="Multivariate Analysis Results.xlsx"){
  #####GENERATING FORMULA#####
  dt<-as.data.frame(dt)
  vi<-dplyr::select(dt,c(vi_name))
  cov<-dplyr::select(dt,con_name,cat_name)
  cov_names<-colnames(cov)

  if(length(group)>1){
    group<-group
  }else{
    group<-dplyr::select(dt,group)
  }

  if(length(ref)!=0){
    group<-relevel(factor(group),ref = ref)
  }else{
    group=factor(group)
  }

  #Create dataset for modeling:
  if(length(con_name)>0&length(cat_name)>0){
    con_cov<-as.data.frame(apply(apply(dplyr::select(dt,con_name),2,as.character),2,as.numeric))
    cat_cov<-as.data.frame(apply(dplyr::select(dt,cat_name), 2, as.factor))
    data<-cbind.data.frame(con_cov,cat_cov,vi,group=group)
  }else{
    if(length(con_name)==0&length(cat_name)>0){
      cat_cov<-as.data.frame(apply(dplyr::select(dt,cat_name), 2, as.factor))
      data<-cbind.data.frame(cat_cov,vi,group=group)
    }else{
      if(length(con_name)>0&length(cat_name)==0){
        con_cov<-as.data.frame(apply(apply(dplyr::select(dt,con_name),2,as.character),2,as.numeric))
        data<-cbind.data.frame(con_cov,vi,group=group)
      }else{
        data<-cbind.data.frame(vi,group=group)
      }
    }
  }

  #Create Formula for model:
  if(length(cov_names)>0){
    form<-formula(paste('y','~','group','+',paste(cov_names,collapse = " + ")))
  }else{
    form<-formula(paste('y','~','group'))
  }
    


  re<-c()
  if(method=="logistic"){
    re2<-c()
    for (i in 1:ncol(vi)){
      y<-factor(vi[,i])
      tmp<-glm(formula = form,data = data,
               family = "binomial")
      coef<-round(summary(tmp)$coefficients,3)
      coef<-cbind.data.frame(c(vi_name[i],rep("",nrow(coef)-1)),
                             rownames(coef),coef)
      k<- exp(cbind(OR = coef(tmp), confint(tmp)))
      wald<-wald.test(b = coef(tmp), Sigma = vcov(tmp), Terms = 2:nlevels(group))
      or<-c("Ref",
            paste(round(k[2:nlevels(group),1],2),
                  " (",
                  round(k[2:nlevels(group),2],2),
                  ", ",
                  round(k[2:nlevels(group),3],2),
                  ")",sep = ""),
            round(wald$result$chi2[3],3))
      re<-rbind.data.frame(re,'',coef)
      nper<-c()
      for (j in 1:nlevels(group)) {
        ftb<-paste(table(y,group)[,j],
                   " (",
                   100*(round(prop.table(table(y,group),2),2)[,j]),
                   ")",sep = "")
        nper<-cbind(nper,ftb)
      }

      OR<-cbind.data.frame(c(vi_name[i],rep("",nlevels(y)-1)),
                           levels(y),
                           nper,
                           rbind(or,matrix("",nrow = nlevels(y)-1,ncol = nlevels(group)+1)))
      re2<-rbind(re2,'',OR)
    }
    re2<-as.matrix(re2)
    rownames(re2)<-as.matrix(re2[,1])
    re2<-re2[,-1]
    colnames(re2)<-c("",
                     levels(group),
                     paste("OR",levels(group), "vs.",ref, "(95% CI)"),
                     "P value")
    colnames(re)<-c("","",colnames(summary(tmp)$coefficients))
    #Add title and footnotes for the  result table:

    l<-list(re2,re)
    #Formating:
    wb<-ExcelTable(table = re2,
                   title ='Table 2. Multivariate Results',
                   caption = paste("Table 2. Multivariate Results. Odds ratios were obtained obtained by logistic regression model adjusting for",paste(cov_names, collapse = ', '),
                                   ". P value were obtained from Wald's test. Frequency were presented as N (%)"),
                   outfile = outfile,
                   openfile = F)
    addWorksheet(wb,"Coefficients & ANOVA")
    writeData(wb,sheet = "Coefficients & ANOVA",re)
    saveWorkbook(wb,outfile,overwrite = T)
    openXL(outfile)
  }else{
    #####For linear regression MODEL####
    re1<-c()
    betas<-list()
    for (i in 1:ncol(vi)) {
      vi[,i]<-as.numeric(as.character(vi[,i]))
      if(log.transform==T){
        y<-log(vi[,i]+1)
      }else{
        y<-vi[,i]
      }
      tmp<-lm(formula = form,data)
      #Coefficients:
      a<-as.matrix(round(coefficients(summary(tmp)),3))
      if(log.transform==T){
        a[,1:2]<-exp(a[,1:2])

      }
      name <- vi_name[i]
      betas[[name]] <- a
      a<-rbind(colnames(a),a)

      #Anova Table:
      #b<-as.matrix(round(anova(tmp),3))
      b<-as.matrix(round(Anova(tmp),3))  #Need 'car' package
      b<-rbind(colnames(b),b)
      #a<-round(coefficients(summary(tmp))[2,],3)
      re<-rbind(re,
                c(colnames(vi)[i],rep("",3)),
                c("Coefficients:",rep("",3)),
                a,
                c("ANOVA table:",rep("",3)),
                b,
                "")
      colnames(re)<-rep("",4)

      #Result table:
      temp1<-a[3:(nlevels(group)+1),]
      if(CI==T){
        ci<-paste('(',round(confint(tmp)[,1],1),', ', round(confint(tmp)[,2],1),')',sep = "")
        civ<-ci[2:nlevels(group)]
        if(nlevels(group)==2){
          evi<-paste(round(as.numeric(temp1[1]),2),civ)
          tb<-c("0",evi,temp1[4])
        }else{
          evi<-paste(round(as.numeric(temp1[,1]),2),civ)
          tb<-c("0",evi,b[2,4])
        }
      }else{
        if(is.vector(temp1)==T){
          temp2<-t(rbind(c(0,"-"), 
                         cbind(paste(round(as.numeric(temp1[1]),2),' ± ',round(as.numeric(temp1[2]),2),sep = ""),
                               temp1[4])))
        }else{
          temp2<-t(rbind(c(0,"-"), 
                         cbind(paste(round(as.numeric(temp1[,1]),2),' ± ',round(as.numeric(temp1[,2]),2),sep = ""),
                               temp1[,4])))
        }
        if(nlevels(group)>2){
          tb<-c(t(temp2[1,]),b[2,4])
        }else{
          tb<-c(t(temp2[1,]),temp1[4])
        }
        }
      tb[2:nlevels(group)]<-paste(tb[2:nlevels(group)],ifelse(as.numeric(a[3:(nlevels(group)+1),4])>0.05,
                                                              "",
                                                              ifelse(as.numeric(a[3:(nlevels(group)+1),4])<=0.05 & as.numeric(a[3:(nlevels(group)+1),4])>0.01, 
                                                                     "*",
                                                                     "**")))
      re1<-rbind(re1,tb)
      }

                                                   
    re1[,ncol(re1)]<-ifelse(as.numeric(re1[,ncol(re1)])>=0.05, 
                    re1[,ncol(re1)],
                    ifelse(as.numeric(re1[,ncol(re1)])<0.05 & as.numeric(re1[,ncol(re1)])>=0.01,
                           paste(re1[,ncol(re1)],"*"),
                           ifelse(as.numeric(re1[,ncol(re1)])<0.01 & as.numeric(re1[,ncol(re1)])>=0.001,
                                  paste(re1[,ncol(re1)],"**"), "<0.001 **")))
                           
    
    rownames(re1)<-colnames(vi)
    colnames(re1)<-c(paste(levels(group), " vs. ",ref," (Mean Difference)",sep = ""),
                     "P value")
    re<-cbind(rownames(re),re)
    l<-list(re1,re)
    #Formating:
    #caption:
    
    if(CI==T){
      if (nlevels(group)>2) {
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ". P Value were obtained from F test. '*':  P (Difference comparing to the reference group) = 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }else{
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ".  '*':  P (Difference comparing to the reference group) < 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }
    }else{
      if (nlevels(group)>2) {
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ". P Value were obtained from F test. '*':  P (Difference comparing to the reference group) = 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }else{
        wb<-ExcelTable(table = re1,
                       title ='Table 2. Multivariate Results',
                       caption = paste("Table 2. Multivariate Results. Adjusted mean differences were presented in Mean (95% CI) and obtained by linear regression model adjusting for",paste(cov_names, collapse = ', '),
                                       ".  '*':  P (Difference comparing to the reference group) < 0.05; '**': P < 0.01"),
                       outfile = outfile,
                       openfile = F)
      }
    }

    addWorksheet(wb,"Coefficients & ANOVA")
    writeData(wb,sheet = "Coefficients & ANOVA",re)
    saveWorkbook(wb,outfile,overwrite = T)
    openXL(outfile)
  }
  return(l)
  }

###########################Post Hoc Analysis Table##################
##Desciption##
#Function for testing of difference between group

#Packages #Packages needed: dplyr, openxlsx

##Arguments##
#vi             #Matrix including variables of interest (continues variables)
#group          #Treatment groups (Exposures). A vector.
#methods        #one of "hsd", "bonf", "lsd", "scheffe", "newmankeuls", defining the method for the pairwise comparisons.
#based.on.aov   #a logical value indicating if only present the post hoc results of the variables that have significant ANOVA result  
#outfile        #Output file location

##Function##
PostHoc<-function(vi, group, methods="hsd", based.on.aov=F , outfile='Post Hoc Analysis.xlsx'){
  tb<-c()
  anova<-c()
  RE<-c()
  vi<-apply(apply(vi,2,as.character), 2, as.numeric)
  for (i in 1:ncol(vi)) {
    test<-cbind.data.frame(vi[,i], group)
    test[,1]<-as.character(as.numeric(test[,1]))
    test.aov<-aov(`vi[, i]`~group,data=test)
    anova<-c(anova,unlist(summary(test.aov))[9])
    re<-PostHocTest(test.aov, adj.methods=methods)$group
    a<-t(PostHocTest(test.aov, adj.methods=methods)$group)
    b<-paste(round(a[1,],2), ' (', round(a[2,],1), ', ', round(a[3,],1), ')', sep = '')
    adj.p<-a[4,]
    b[which(adj.p<0.05)]<-paste(b[which(adj.p<0.05)], '*')
    tb<-rbind(tb,b)
    
    rn<-c(colnames(vi)[i],paste("          ",rownames(re)))
    re<-rbind("",round(re,2))
    rownames(re)<-rn
    RE<-rbind(RE,re,"")
  }
  rownames(tb)<-colnames(vi)
  colnames(tb)<-paste(colnames(a), " (95% CI)", sep = "")
  colnames(RE)<-c("Mean Difference", "Lower 95% CI", "Higher 95% CI", "P Values")
  
  if(based.on.aov==T){
    tb<-tb[which(anova<0.05),]
  }
  
  ExcelTable(RE, 
             title="Post Hoc Analysis",
             caption=paste( methods, "method was used to correct for multiple comparison"),
             outfile=outfile,
             openfile=T)
  return(RE)
}




###########################SF-12 Calculator##################
SF12<-function(data){
  data<-as.data.frame(apply(data, 2, as.numeric))
  sf12<-c()
  sf12$PF <- data[,2] + data[,3]
  sf12$RP <- data[,4] + data[,5]
  sf12$BP <- data[,8]
  sf12$GH <- data[,1]
  sf12$VT <- data[,10]
  sf12$SF <- data[,12]
  sf12$RE <- data[,6] + data[,7]
  sf12$MH <- data[,9] + data[,11]
  sf12<-as.data.frame(sf12)
  
  sf12$TSPF<-(sf12$PF-2)/4*100
  for (i in 2:(ncol(sf12)-1)) {
    if(i %in% c(3,4,5,6)){
      sf12[,i+8]<-(sf12[,i]-1)/4*100
    }else{
      sf12[,i+8]<-(sf12[,i]-2)/8*100
    }
  }
  sf12[,17]=(sf12[,9]-81.18122)/29.10558
  sf12[,18]=(sf12[,10]-80.52856)/27.13526
  sf12[,19]=(sf12[,11]-81.74015)/24.53019
  sf12[,20]=(sf12[,12]-72.19795)/23.19041
  sf12[,21]=(sf12[,13]-55.5909)/24.8438
  sf12[,22]=(sf12[,14]-83.73973)/24.75775
  sf12[,23]=(sf12[,15]-86.41051)/22.35543
  sf12[,24]=(sf12[,16]-70.18217)/20.50597
  
  for (i in 17:24) {
    sf12[,i+8]=(sf12[,i]*10)+50
  }
  sf12[,33]=(sf12[,17]*0.42402)+(sf12[,18]*0.35119)+(sf12[,19]*0.31754)+(sf12[,20]*0.24954)+(sf12[,21]*0.02877)+(sf12[,22]*-0.00753)+(sf12[,23]*-0.19206)+(sf12[,24]*-0.22069)
  sf12[,34]=(sf12[,17]*-0.22999)+(sf12[,18]*-0.12329)+(sf12[,19]*-0.09731)+(sf12[,20]*-0.01571)+(sf12[,21]*0.23534)+(sf12[,22]*0.26876)+(sf12[,23]*0.43407)+(sf12[,24]*0.48581)
  
  sf12[,35]=50+(sf12[,33]*10)
  sf12[,36]=50+(sf12[,34]*10)
  
  colnames(sf12)<-c('raw score Physical Functioning (PF)',	'raw score Role Physical (RP)',	'raw score Bodily Pain (BP)',	
                    'raw score General Health (GH)',	'raw score Vitality (VT)',	'raw sore Social Functioning (SF)',	
                    'raw score Role Emotional (RE)',	'raw score Mental Health (MH)',	'transformed score PF',	'transformed score RP',	
                    'transformed score BP',	'transformed score GH',	'transformed score VT',	'transformed score SF',	
                    'transformed score RE',	'transformed score MH',	'PF z-score',	'RP z-score',	'BP z-score',	'GH z-score',	
                    'VT z-score',	'SF z-score',	'RE z-score',	'MH z-score',	'PF norm based',	'RP norm based ',	'BP norm based ',	
                    'GH norm based ',	'VT norm based ',	'SF norm based ',	'RE norm based ',	'MH norm based ',	'Aggregate PHYSICAL',
                    'Aggregate MENTAL',	'SF12_PCS',	'SF12_MCS')
  return(sf12)
  
}



###########################Correlation Matrix##################
corr.table<-function(col, row, method="pearson", outfile="Correlation Tables.xlsx"){
  col<-apply(apply(col,2,as.character), 2, as.numeric)
  row<-apply(apply(row,2,as.character), 2, as.numeric)
  col<-as.data.frame(col)
  row<-as.data.frame(row)
  
  library("Hmisc")
  cor<-rcorr(x=as.matrix(col),y=as.matrix(row), type=method)
  
  R<-apply(cor$r[(ncol(col)+1):nrow(cor$r),1:ncol(col)], 2, round, digit=2)
  P<-apply(cor$P[(ncol(col)+1):nrow(cor$r),1:ncol(col)], 2, round, digit=3)
  
  CM<-rbind(R,c(""),P)
  
  RP<-c()
  rowname<-c()
  for (i in 1:ncol(row)) {
    r<-R[i,]
    p<-P[i,]
    r[which(p<0.05)]<-paste(r[which(p<0.05)],"*", sep = "")
    rp<-cbind(c("Correlation Coefficient", "P value"),rbind(r,p))
    RP<-rbind(RP,"",rp)
    rowname<-c(rowname,c(colnames(row)[i],"",""))
  }
  rownames(RP)<-rowname
  l<-list(RP, CM)

  wb<-ExcelTable(table = RP, 
                 title = "Correlation Matrix",
                 caption = paste("Correlation Table. Values were ",
                                 method,
                                 " correlation coefficients. '*' P< 0.05",
                                 sep = ""
                                 ) ,
                 outfile = outfile,
                 openfile = F)
  
  # addWorksheet(wb,"Correlation Matrix")
  # writeData(wb,sheet = "Correlation Matrix", CM)
  # saveWorkbook(wb,outfile,overwrite = T)
  openXL(outfile)
  return(l)
}
#####
