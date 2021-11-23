import pandas as pd
import dateutil


def adjust_month(s, delta=0, strft_fmt="%b %Y", raise_exception = True):
    
    '''
    adjust month by number of months
    
    s => can be a date, or a string e.g. "feb 2020"
    delta => number of months (negative to reverse)
    strft_fmt = None -> return date
                else -> return string formatted date
                
    Example:
    adjust_month('jan 2020', +1) => "Feb 2020"
    adjust_month('jan 2020', -3) => 'Oct 2019'
    adjust_month('jan 2020', +1, strft_fmt=None) = datetime.date(2019, 10, 1)
    
    #20210120 - added raise_exception to allow error to be switched off.
                useful when using this function to screen through the 
                list of possible months and it is expected that some 
                are not months.
    '''
    
    # Convert to date
    try:
        d = pd.to_datetime(s).date()
    except:
        
        if raise_exception is True:
            error = "Cannot parse date for '%s'." % s
            logger.error(error)
            raise Exception (error)
        else:
            return None
    
    # Set time unit
    month_unit = dateutil.relativedelta.relativedelta(months=delta)
    
    # return
    adj_d = d + month_unit
    
    if strft_fmt is None:
        return adj_d
    else:
        return adj_d.strftime(strft_fmt)