
# Libraries Needed:
library(readxl)
library(car)
library(forecast)
library(lubridate)
library(ggpubr)
library(tidyverse)
library(ggplot2)
library(anytime)
library(reshape2)
library(prophet)
'%notin%'=Negate('%in%')


# filepath
file_path=paste('C:/Users/Muhammad Hussain/Desktop','Arconic Data 1.xlsx',sep="/")

# Reading Data Files:

# Reading Historical Data:
hist_data=read_excel(file_path,sheet=1)

# Open Order Data:
open_data=read_excel(file_path,sheet=2)

hist_data <- hist_data %>% select(`Item Group (Product Family)`,`Order Date`,`Item #`,`Quantity (PCS)`,`Sales Order #`)
colnames(hist_data)<- c('item_group','order_date','item_num','quantity','sales_order')

open_data <- open_data %>% select(`Item Group (Product Family)`,`Order Date`,`Item #`,`Quantity (PCS)`,`Sales Order #`)
colnames(open_data)<- c('item_group','order_date','item_num','quantity','sales_order')

# Removing Duplicates:
final_data=hist_data %>% rbind(open_data)
dup_indexes=which(duplicated(final_data)==TRUE)
final_data=final_data[-dup_indexes,]


final_data$quantity=as.numeric(final_data$quantity)
final_data=final_data %>% group_by(item_group,item_num,order_date) %>% summarize(total_orders=sum(quantity)) %>% 
arrange(item_group,item_num,order_date)

final_data=final_data %>% mutate(week=week(order_date),month=month(order_date))


final_data=final_data %>% mutate(datu=format(order_date,"%Y-%m")) %>%
group_by(item_group,item_num,datu) %>% summarise(total_orders=sum(total_orders)) %>%
arrange(item_group,item_num,datu)



items=final_data %>%  group_by(item_group,item_num) %>% count() %>% filter(n>12) %>% select(item_group,item_num)
finalu=final_data %>% filter(item_group %in% items$group | item_num %in% items$item_num) %>% arrange(item_group,item_num,datu)

# showing how many items are there in each group
options(repr.plot.width=17, repr.plot.height=7)
plot_data=finalu %>% group_by(item_group) %>% summarise(num_items=n_distinct(item_num)) %>% arrange(num_items)
ggplot(plot_data,aes(x=reorder(item_group,-num_items),y=num_items,fill=item_group))+geom_bar(stat='Identity') +
theme(axis.text.x = element_text(angle = 60)) + labs(x='Item Group (Family of Items)',y='Number of Unique items in Family')

# Selecting only those product items for which we have data till at least July-2018, so that we can do forecasting for the rest of the year.
ds=finalu %>% group_by(item_group,item_num) %>% summarize(datu=last(datu))
ds=ds[which(anydate(ds$datu)>='2018-07-01'),] %>% select(item_group,item_num)
finalu=finalu %>% inner_join(ds,by=c("item_group","item_num"))

dats=finalu %>% filter(item_group=='11282')
options(repr.plot.width=30, repr.plot.height=15)
ggplot(dats,aes(x=datu,y=total_orders,group=1))+geom_point(color='steelblue',size=2)+
geom_line(color='steelblue',size=2)+
labs(x='Date - Year & Month',y='Order Quantity')+ theme(axis.text.x = element_text(angle = 60))+
facet_wrap(item_num~.)


# Time Series Object:
time_ser <- function(data){
    start_date=data$datu[1]
    stm=month(start_date)
    sty=year(start_date)
    
    end=data$datu[nrow(data)]
    edm=month(end)
    edy=year(end)
    time_ser_obj=ts(data$total_orders,start=c(sty,stm),end=c(edy,edm),frequency=12)
    return(time_ser_obj)
}

# SSE Error Calculation:
sse_calc <- function(actual,forecast){
    sse=sqrt(sum((actual-forecast)^2)/length(forecast))
    return(sse)
}

# Moving Average Method:
ma_method <- function(series){
    main_error=c()
    for (order in c(1:7)){
        model=ma(series,order)
        for (i in c(1:length(series))){
            if(is.na(model[i])==TRUE){
                model[i]=series[i]
            }
        }
        err=sse_calc(series,model)
        main_error=c(main_error,err)
    }
    w=which.min(main_error)
    return(w)
}


# Complex Arima Function:
complex_arima<- function(series){
    lamd=c(0,1/2,1,2)
    err=c()
    error_happened=c()
    for (lam in lamd){
        a=tryCatch({
            error=0
            auto.arima(ts_dat,seasonal=TRUE,lambda=lam)
            
        },error=function(e){
            error=1
            auto.arima(ts_dat)
        },finally={
            error=1
            ets(ts_dat)
        })
        error_happened=c(error_happened,error)
        s=sse_calc(series,a$fitted)
        err=c(err,s)
    }
    w=which.min(err)
    if (w==1){
        return(c(0,error_happened[1]))
    }else if (w==2){
        return(c(1/2,error_happened[2]))
    }else if (w==3){
        return(c(1,error_happened[3]))
    }else{
        return(c(2,error_happened[4]))
    }
}


# Missing Value Imputation Function:
miss_val_imputation <- function(data){
    
    start_date=data$datu[1]
    end_date=data$datu[nrow(data)]

    sequ=seq(from=start_date,to= end_date,by='month')
    notins=which(sequ %notin% data$datu)

    for(x in notins){
        if(x>12){
            data=data %>% add_row(datu=sequ[x],total_orders=data$total_orders[x-12])
        }else{
            data=data %>% add_row(datu=sequ[x],total_orders=0)
        }
    }
    data=data[!duplicated(data),]
    return(data)            
}

# Prophet Time Series Object:
prophet_tser <- function(data){
    ds=data$datu
    y=data$total_orders
    main_datu=data.frame(ds,y)
    return(main_datu)
}



# Advanced Arima Function:
advanced_arima <- function(data,order,lamda){
    erru=c()
    rownames=lamda
    colnames=order
    if (length(order)==1){
        for (lamda in lamda){
            a1=auto.arima(data,seasonal=TRUE,lambda=lamda)
            err=sse_calc(data,a1$fitted)
            erru=c(erru,err)
        }
        m=matrix(erru,nrow=1,ncol=length(lamda))
        colnames(m)<- lamda
        w=which.min(m)
        lamdu_final=as.double(colnames(m)[w])
        return(lamdu_final)
        
    }else{
        main_err=c()
        for (lamda in lamda){
            err_1=c()
            for (order in order){
                model=auto.arima(data,xreg=fourier(data,K=order),lambda=lamda,seasonal=FALSE)
                err=sse_calc(data,model$fitted)
                err_1=c(err_1,err)
            }
            
            main_err=rbind(main_err,err_1)
        }
        m=as.matrix(main_err)
        rownames(m)<- rownames
        colnames(m)<- colnames
        w=c(which(m==min(m),arr.ind=TRUE))

        best_lamda=as.double(rownames(m)[w[1]])
        best_order=as.numeric(colnames(m)[w[2]])
        
        return(c(best_lamda,best_order))
    }
}

# Main Function:
mains_func <- function(o_dat,p_dat,order,lamda,type){
    # Prophet Data = p_dat
    # original Data = o_dat
    error=c()
    
    if (type=='additive'){
        
        advanced_params=tryCatch({
          error=0
          advanced_arima(o_dat,order,lamda)},
                                error=function(e){
                                  error=1
                                  1
                                },
          finally={
            error=1
            1
          })
        if (error==1){
            a1=ets(o_dat)
        }else if(error==0){
            a1=auto.arima(o_dat,lambda=advanced_params,seasonal=TRUE)
        }else{
          a1=ets(o_dat)
        }
        error=c(error,sse_calc(o_dat,a1$fitted))
        
        
        # Tbats:
        a2=tbats(o_dat)
        error=c(error,sse_calc(o_dat,a2$fitted))
        
        
        # Prophet:
        a3=prophet(p_dat,seasonality.mode=type)
        future=make_future_dataframe(a3,periods=6,freq='month')
        prediction=predict(a3,future)
        
        
        forecast_prophet=prediction %>% select(yhat) %>% tail(6)
        fitted_prophet=prediction %>% select(yhat) %>% head(nrow(p_dat))
        
        error=c(error,sse_calc(o_dat,fitted_prophet$yhat))
        
        mins=which.min(error)
        if (mins==1){
            forecast=forecast(a1,h=6)
            return(forecast$mean)
        }else if (mins==2){
            forecast=forecast(a2,h=6)
            return(forecast$mean)
        }else{
            forecast=forecast_prophet$yhat
            return(forecast)
        }
        
    }else{
        advanced_params=tryCatch({
          error=0
          advanced_arima(o_dat,order,lamda)},
                                 error=function(e){
                                   error=1
                                   c(1,0)
                                 },
          finally={error=1
          c(1,0)})
        
        lamda=advanced_params[1]
        order=advanced_params[2]
        
        # Arima:
        if(error==1){
          a1=auto.arima(o_dat)
        }else{
          a1=auto.arima(o_dat,xreg=fourier(o_dat,K=order),lambda=lamda,seasonal=FALSE)
        }
        error=c(error,sse_calc(o_dat,a1$fitted))
        
        
        # Tbats:
        a2=tbats(o_dat)
        error=c(error,sse_calc(o_dat,a2$fitted))
        
        
        # Prophet:
        a3=prophet(p_dat,seasonality.mode=type)
        future=make_future_dataframe(a3,periods=6,freq='month')
        prediction=predict(a3,future)
        
        
        forecast_prophet=prediction %>% select(yhat) %>% tail(6)
        fitted_prophet=prediction %>% select(yhat) %>% head(nrow(p_dat))
        
        error=c(error,sse_calc(o_dat,fitted_prophet$yhat))
        
        mins=which.min(error)
        if (mins==1){
            if(error==1){
              forecast=forecast(a1,h=6)
            }else{
              forecast=forecast(a1, xreg=fourier(o_dat, K=order, h=6))
            }
            
            return(forecast$mean)
        }else if (mins==2){
            forecast=forecast(a2,h=6)
            return(forecast$mean)
        }else{
            forecast=forecast_prophet$yhat
            return(forecast)
        }
        
    }
    
}


# Main Forecasting Tool:

final_forecast_dat=c()
# c(unique(finalu$item_group)
for (item_group_zz in c(unique(finalu$item_group))){
  print(item_group_zz)
  dats=finalu %>% filter(item_group==item_group_zz)
  fou=c()
  for (item in c(unique(dats$item_num))){
    main_data=dats %>% filter(item_num==item) %>%
      mutate(datu=anydate(datu)) %>% arrange(datu)
    
    main_data=main_data[,c('datu','total_orders')]
    main_data=miss_val_imputation(main_data)
    main_data=main_data %>% arrange(datu)
    
    data_dim=dim(main_data)[1]
    
    # time Series Object Creation:
    ts_dat=time_ser(main_data)
    
    # checking How many obs are zeroes:
    zeros=which(ts_dat==0)
    
    if (data_dim>36 & length(zeros)<=3){
      
      prophet_ts=prophet_tser(main_data)
      
      # checking Seasonality:
      add=sum((Acf(decompose(ts_dat,type='additive')$random)$acf)^2)
      mult=sum((Acf(decompose(ts_dat,type='multiplicative')$random)$acf)^2)
      wik=which.min(c(add,mult))
      
      if (wik==1){
        type='additive'
        order=0
        lamda=c(0,1,1/2,1/3)
        forecast=mains_func(ts_dat,prophet_ts,order,lamda,type)
        
      }else{
        type='multiplicative'
        order=c(1,2,3,4,5)
        lamda=c(0,1,1/2,1/3)
        forecast=mains_func(ts_dat,prophet_ts,order,lamda,type)
      }
    }
    else if (data_dim<3){
      
      forecast=rep(mean(main_data$total_orders),6)
      
    }
    else if (dim(main_data)[1]<=15 & length(zeros<=2)){
      exp_smooth=ets(ts_dat)
      moving_avg=ma(ts_dat,ma_method(ts_dat))
      for (i in c(1:length(ts_dat))){
        if(is.na(moving_avg[i])==TRUE){
          moving_avg[i]=series[i]
        }
      }
      arima=auto.arima(ts_dat)
      s1=exp_smooth$SSE
      s2=sse_calc(ts_dat,moving_avg)
      s3=sse_calc(ts_dat,arima$fitted)
      s=c(s1,s2,s3)
      w=which.min(s)
      if (w==1){
        forecast=forecast(exp_smooth,h=6)
        forecast=forecast$mean
      } else if (w==2){
        forecast=forecast(moving_avg,h=6)
        forecast=forecast$mean
      }else{
        forecast=forecast(arima,h=6)
        forecast=forecast$mean
      }
    }else{
      exp_smooth=hw(ts_dat)
      comp_return=complex_arima(ts_dat)
      if(comp_return[2]==1){
        arima=auto.arima(ts_dat)
      }else{
        lamdu=comp_return[1]
        arima=auto.arima(ts_dat,seasonal=TRUE,lambda=lamdu)
      }
      tbatu=tbats(ts_dat,lambda=lamda)
      s1=exp_smooth$SSE
      s2=sse_calc(ts_dat,arima$fitted)
      s3=sse_calc(ts_dat,tbatu$fitted)
      s=c(s1,s2,s3)
      w=which.min(s)
      if (w==1){
        forecast=forecast(exp_smooth,h=6)
        forecast=forecast$mean
      } else if (w==2){
        forecast=forecast(arima,h=6)
        forecast=forecast$mean
      }else{
        forecast=forecast(tbatu,h=6)
        forecast=forecast$mean
      }
    }
    
    fo=c(item_group_zz,item,forecast)
    fou=rbind(fou,fo)
  }
  final_forecast_dat=rbind(final_forecast_dat,fou)
}

# Making a copy
final_forecast_copy=final_forecast_dat


# Final Manipulations:
dates=seq.Date(as.Date('2018-10-01'),length.out = 6,by='month')

colnames(final_forecast_dat)<- c('item_group','item_num',sapply(dates,as.character))

final_forecast_dat=as.data.frame(final_forecast_dat)
final_forecast_dat[,3:8]=sapply(final_forecast_dat[,3:8],function(x) as.numeric(as.character(x)))


final_forecast_dat=reshape2::melt(final_forecast_dat,id.vars=c('item_group','item_num'),variable.name='month',value.name='forecast') %>%
  mutate(forecast=round(forecast)) %>% arrange(item_group,item_num,month)


final_forecast_dat$month=format.Date(final_forecast_dat$month,format='%Y-%m')

View(final_forecast_dat)

# Writing as CSV file:
write_excel_csv(final_forecast_dat,'C:/Users/Muhammad Hussain/Desktop/Final_Arconic_Forecast_newewst.csv')


