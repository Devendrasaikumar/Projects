library(glmnet)
library(data.table)
library(readxl)

# Importing Dataset

my_data <- read_xlsx("D:/Sai/APP/APB Final Data Merging_V2_06.12.2024 (Overall Data).xlsx", sheet = "APB Final Data (For Model)", col_names = TRUE)
str(my_data)		#structure of the data
names(my_data)
dim(my_data)		#dimension of the data
colSums(Filter(is.numeric, my_data))      # Column Total

sales_data <- my_data
names(sales_data)

# Define Indep. variable
ind_var<-c(9:117)
noofIndvar <- length(ind_var)
noofIndvar

# Define Dep. variable
dep_var <- c("HNK_ORG_SALES")

# creating dummy dataframe for merging purpose
a <- data.frame(row.names=0:noofIndvar) 
b_trainerr <- data.frame(row.names="Train_Error")
c_testerr <- data.frame(row.names="Test_Error")
d_trainmape <- data.frame(row.names="Train_Mape")
e_testmape <- data.frame(row.names="Test_Mape")

#looping
 for(i in 1:1000) 
  set.seed(i)
  
  #setting random sample for test and train data
  dt = sort(sample(nrow(sales_data), nrow(sales_data)*.9))
  dt

  #Splitting dataset into train(90%) and test(10%).
  train_sales<-sales_data[dt,]
  test_sales <-sales_data[-dt,]
  
  #Define Explanatory and Response Variable for train and test
  train_depvar<-train_sales[,dep_var]
  train_depvar<-as.matrix(train_depvar)
  
  test_depvar <- test_sales[,dep_var]
  test_depvar <- as.matrix(test_depvar)

  train_indvar<-train_sales[,ind_var]
  train_indvar<-as.matrix(train_indvar)

  test_indvar<-test_sales[,ind_var]
  test_indvar<-as.matrix(test_indvar)

  
  #------------------------------------------------------------------------#
  
  #--> Model building
  # alpha = 1 is the lasso penalty,  alpha = 0 the ridge penalty.
  
  fit<-glmnet(x=train_indvar ,y=train_depvar, alpha = 1, lambda = 0, lower.limits = 0, intercept= TRUE, standardize = F)
  coef(fit,s=c(0.01))					#to get the Co-efficient
  coeff=as.matrix(coef(fit,s=c(0.01)))
  coeff_df=as.data.frame(coeff)  
  
  #Merging the coefficient for each iteration
  a <- cbind(a, coeff_df)  
  rownames(a)=rownames(coeff_df)  
  names(a)[ncol(a)] <- paste0("Coeff_Run_", i)  
  
  #Prediction of Train and Test
  train_pred<-data.frame(predict(fit,train_indvar,s=c(0.01)))
  test_pred<-data.frame(predict(fit,test_indvar,s=c(0.01)))

  #Calculation for absolute error
  train_pred <- cbind(train_pred, train_depvar, (abs((train_pred/train_depvar)-1)))
  test_pred <- cbind(test_pred, test_depvar, (abs((test_pred/test_depvar)-1)))  

  #Defining the column names for prediction
  train_pred <- setNames(train_pred,c("Predicted","Actual","ABS_Error"))
  test_pred <- setNames(test_pred,c("Predicted","Actual","ABS_Error"))  

  train_pred$ABS_Error <- ifelse(train_pred$ABS_Error == "Inf",0 ,train_pred$ABS_Error)
  test_pred$ABS_Error <- ifelse(test_pred$ABS_Error == "Inf",0 ,test_pred$ABS_Error)  

  #calculation for ERROR & MAPE
  train_pred_total <- colSums(train_pred[1])
  train_act_total <- colSums(train_pred[2])
  train_error_pct <- ((train_pred_total/train_act_total)-1)   
 
  b_trainerr <- cbind(b_trainerr, train_error_pct)

  test_pred_total <- colSums(test_pred[1])
  test_act_total <- colSums(test_pred[2])
  test_error_pct <- ((test_pred_total/test_act_total)-1)  

  c_testerr <- cbind(c_testerr, test_error_pct)  

  train_mape <- colMeans(train_pred[3])
  test_mape <- colMeans(test_pred[3])  

  d_trainmape <- cbind(d_trainmape, train_mape)
  e_testmape <- cbind(e_testmape, test_mape)  

  # Merging all ERROR & MAPE Values
  names(b_trainerr)[ncol(b_trainerr)]   <- paste0("Run_", i)
  names(c_testerr)[ncol(c_testerr)]     <- paste0("Run_", i)
  names(d_trainmape)[ncol(d_trainmape)] <- paste0("Run_", i)
  names(e_testmape)[ncol(e_testmape)]   <- paste0("Run_", i)  

  overall_summary <- rbind(b_trainerr,c_testerr,d_trainmape,e_testmape)  

  # ---------------------------------------------------------------------------------
  
  #Predict the ERROR and MAPE for Average coefficient
  Reqdata_indvar <- sales_data[,ind_var]
  Reqdata_depvar <- sales_data[,dep_var]
  Reqdata_depvar <- as.data.frame(Reqdata_depvar)
  Reqdata_depvar <- setNames(Reqdata_depvar,c("Actual_Value"))

  PredData <- as.data.frame(t(as.matrix(Reqdata_indvar)))
  noofvar <- dim(PredData)
  noofvar <- noofvar[2]  

  PredData <- rbind(Cons = 1,PredData)  

  coeffvarname <- as.data.frame(dimnames(a)[[1]])
  avgcoeff <- cbind(coeffvarname , as.data.frame(a))
  
  f <- data.frame(row.names="Pred_Value")
  for(i in 1:noofvar) {
    x <- colSums(data.frame(PredData[,c(i)]) * data.frame(avgcoeff[,c(2)]))
    f <- cbind(f, x)
  }  

  #f
  zpred <- as.data.frame(t(as.matrix(f)))

  zpred <- cbind(Reqdata_indvar, Reqdata_depvar, zpred)  

  zpred$error_pct <- ((zpred$Pred_Value / zpred$Actual_Value)-1)
  zpred$error_pct <- ifelse(zpred$error_pct == "Inf",0 ,zpred$error_pct)  

  zpred$ABSerror_pct <- abs(zpred$error_pct)  

  #calculation for ERROR & MAPE
  zpred_total <- as.data.frame(c(sum(zpred$Pred_Value)))
  zact_total <- as.data.frame(c(sum(zpred$Actual_Value)))
  zerror_pct <- as.data.frame(((zpred_total/zact_total)-1))
  zmape <- as.data.frame(c(mean(zpred$ABSerror_pct)))  

  finalreport <- cbind(zpred_total, zact_total,zerror_pct,zmape) 
  finalreport <- setNames(finalreport,c("Pred_Value", "Actual_Value" , "ErrorPCT", "MAPE"))
  finalreport

  