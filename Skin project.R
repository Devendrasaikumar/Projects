library(readxl)
d1=read_excel("C:/Users/Admin/Downloads/Demo Data - 20 Feb 2023.xlsx",sheet = "R Data")
d1
dim(d1)
colnames(d1)
Rightmean=colMeans(d1[,c("Visit 1 - Baseline","Visit 1 - Timm","Visit 2","Visit 3")])
Leftmean=colMeans(d1[,c( "Visit 1 - Baseline_L", "Visit 1 - Timm_L","Visit 2_L","Visit 3_L")])
Rightmean
Leftmean
x=Rightmean
x
y=Leftmean
y

#### Line Plot ####
plot(x,type = "o",ylab = "Brightness",main = "Line Graph",col="blue",lwd=3)
plot(y,type = "o",ylab = "Brightness",main = "Line Graph",col="blue",lwd=3)
lines(x,type = "o",col="green",lwd=2)

#### Using Bar grpah to compare left and right #####
# Right ##
q=d1$`Visit 1 - Baseline`
mean(q)
w=d1$`Visit 1 - Timm`
mean(w)
e=d1$`Visit 2`
mean(e)
r=d1$`Visit 3`
mean(r)
Rmean=c(mean(q),mean(w),mean(e),mean(r))
X=Rmean
X

# Left ##
t=d1$`Visit 1 - Baseline_L`
mean(t)
y=d1$`Visit 1 - Timm_L`
mean(y)
u=d1$`Visit 2_L`
mean(u)
i=d1$`Visit 3_L`
mean(i)
Lmean=c(mean(t),mean(y),mean(u),mean(i))
Y=Lmean
Y
#### Comparing Bar Graph using their Means #####
X <- c(52.93053, 54.04947, 54.38026, 55.37105)
Y <- c(52.83000, 53.75026, 54.35368, 55.33447)
categories <- c("A", "B", "C", "D")


# Create the first bar plot for X
RX=barplot(
  X,
  beside = TRUE,
  ylim = c(0, max(c(X, Y))),
  names.arg = categories,
  xlab = "Categories",
  ylab = "Values",
  main = "Right V/S Left",
  col = "red",width = 0.2
 
)
# Create the second bar plot for Y, adjust the positions of the bars
LX=barplot(
  Y,
  beside = TRUE,
  add = TRUE,  # Add bars to existing plot
  col = "black",
  width =0.2,
  space = 0.25 # Adjust the position of the bars
  )

# Add legend
legend("bottomleft", legend = c("X", "Y"), fill = c("red", "black"))


#### Using Advanced Comparing ###

data <- data.frame(
  categories = c("A","B","C","D"),
  X = c(52.93053, 54.04947, 54.38026, 55.37105),
  Y = c(52.83000, 53.75026, 54.35368, 55.33447)
)
data
library(plotly)
library(ggplot2)
fig <- plot_ly(data, x =categories , y = X, type = 'bar', name = 'Right_Cheek')
fig <- fig %>% add_trace(y = Y, name = 'Left_Cheek')

fig



s##### Finding Normal Dist for Right skin #####
R=c(d1$`Visit 1 - Baseline`,d1$`Visit 1 - Timm`,d1$`Visit 2`,d1$`Visit 3`)
R
data.frame(R)
mean_valueR=mean(R)
mean_valueR
##### Finding Standard Deviation for Right skin ####
std_R=sd(R)
std_R
##### Finding Normal Distribution  for Right skin ####
nd_R=dnorm(R,mean =mean_valueR,sd=std_R )
nd_R

plot(R,nd_R,main = "Normal Dist R",col = "Blue")


##### Finding Normal Dist for left skin #####
L=c(d1$`Visit 1 - Baseline_L`,d1$`Visit 1 - Timm_L`,d1$`Visit 2_L`,d1$`Visit 3_L`)
L
data.frame(L)
mean_valueL=mean(L)
##### Finding Standard Deviation for left skin ####
std_L=sd(L)
std_L
##### Finding Normal Distribution  for left skin ####
nd_L=dnorm(L,mean_valueL,std_L)
nd_L

plot(L,nd_L,main = "Normal Dist L",col = "Blue")


