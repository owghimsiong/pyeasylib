import pandas as pd
import dateutil
import calendar
import numpy as np
import logging

# Set logger
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# CALENDAR ABBR TO MONTH
month_abbr_to_number = {
    calendar.month_abbr[i] : i
    for i in range(1, 13)
    }
MONTH_ABBR_TO_NUMBER = month_abbr_to_number
MONTH_NAME_TO_NUMBER = {
    calendar.month_name[i] : i
    for i in range(1, 13)
    }

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
    
    
def ensure_correct_date_format(date_series, date_format = 'dmy', 
                               delimiter = None):
    '''
    This method checks for date format.
    
    date_format is assumed to be either:
    I)    dmy: dd.mm.yyyy
    II)   mdy: mm.dd.yyyy
    III)  ymd: yyyy.mm.dd
    
    where:
        d is between 1 to 31
        m is between 1 to 12, Jan to Dec or January to December
        y is between 0 to 9999
        . is the delimiter specified or [- / . <space>] if not specified
    '''
    
    MONTH_ABBR = list(MONTH_ABBR_TO_NUMBER.keys())
    MONTH_NAME = list(MONTH_NAME_TO_NUMBER.keys())
    regex_mth = '|'.join(MONTH_ABBR) + '|' + '|'.join(MONTH_NAME)
    
    if date_series.dtype == object:
        
        # Drop na and dups
        unique_date_series = date_series.dropna().drop_duplicates()
        unique_date_series.index = unique_date_series.values
        
        # Split into 5 columns
        # Y/M/D/B/b smth Y/M/D/B/b smth Y/M/D/B/b
        if delimiter is None:
            regex_pat = (
                r'(?i)(' + regex_mth +
                r'|[0-9]{1,4})[\-/. ]?(' + regex_mth +
                r'|[0-9]{1,4})[\-/. ]?([0-9]{1,4})'
                )
            
            yr_mth_day_df = \
                unique_date_series.astype(str).str.extract(regex_pat)
        
        else:
            yr_mth_day_df = \
                unique_date_series.str.split(delimiter, expand = True)
            yr_mth_day_df[2] = yr_mth_day_df[2].str.extract('^(\d{1,4})')
                
        if date_format.strip().lower() == 'ymd':
            yr_mth_day_df.rename(columns = {0:'yr', 1:'mth', 2:'day'},
                                 inplace = True, errors = 'raise')
        elif date_format.strip().lower() =='mdy':
            yr_mth_day_df.rename(columns = {0:'mth', 1:'day', 2:'yr'},
                                 inplace = True, errors = 'raise')
        elif date_format.strip().lower() == 'dmy':
            yr_mth_day_df.rename(columns = {0:'day', 1:'mth', 2:'yr'},
                                 inplace = True, errors = 'raise')
        else:
            msg = ('date_format is not ["dmy", "mdy", "dmy"], where '
                   'd = Day, m = Month, y = Year.'
                   )
            logger.error(msg)
            raise NotImplementedError (msg)
        
        # convert to month name/abbr to number
        yr_mth_day_df['mth'] = yr_mth_day_df['mth'].str.title()
        all_month_name_ind = yr_mth_day_df['mth'].isin(MONTH_NAME)
        all_month_abbr_ind =yr_mth_day_df['mth'].isin(MONTH_ABBR)
        
        if all_month_abbr_ind.any():
            yr_mth_day_df['mth'] = yr_mth_day_df['mth'].replace(MONTH_ABBR_TO_NUMBER)
        
        elif all_month_name_ind.any():
            yr_mth_day_df['mth'] = yr_mth_day_df['mth'].replace(MONTH_NAME_TO_NUMBER)
        
        
        day_is_mth_ind = yr_mth_day_df['day'].isin(MONTH_NAME + MONTH_ABBR)
        col_id = yr_mth_day_df.columns.get_loc('day')
        
        if day_is_mth_ind.all() and col_id==1:
            msg = (f'date_format = "{date_format}" is incorrect. It should be '
                   'either ["dmy", "ymd"].'
                   )
            logger.error(msg)
            raise Exception (msg)
        
        elif day_is_mth_ind.all() and col_id==0:
            msg = (f'date_format = "{date_format}" is incorrect. It should be '
                   '["mdy"].'
                   )
            logger.error(msg)
            raise Exception (msg)
            
        
        # convert to int
        yr_mth_day_df = yr_mth_day_df.apply(pd.to_numeric, errors ='raise') 
        
        yr_mth_day_df['Day Criterion'] = \
            (yr_mth_day_df['day']>=1) & (yr_mth_day_df['day']<=31)
        
        yr_mth_day_df['Month Criterion'] = \
            ((yr_mth_day_df['mth']>=1) & (yr_mth_day_df['mth']<=12))
            
        yr_mth_day_df['Year Criterion'] = (
            ((yr_mth_day_df['yr']>=0) & (yr_mth_day_df['yr']<=9999))
            )
        yr_mth_day_df['All Criterion'] = (
            yr_mth_day_df[['Day Criterion','Month Criterion','Year Criterion']]
            .all(axis=1))
        
        if not yr_mth_day_df['All Criterion'].all():
            msg = ('Date series is not "{}", where '
                   'd = Day, m = Month, y = Year.'
                   .format(date_format.lower().strip()))
            logger.error(msg)
            raise Exception(msg)
    elif pd.api.types.is_datetime64_any_dtype(date_series):
        msg = ('Already a datetime format').format(date_series.name)
        logger.info(msg)
    else:
        msg = ("Invalid input type: {}. Please convert to pandas.Series."
               .format(date_series.dtype)
               )
        logger.error(msg)
        raise NotImplementedError(msg)


def compare_two_months(first_month, second_month):
    '''
    Return:
        -1: if first_month is earlier than second_month;
         0: if first_month is equal to second_month;
         1: if first_month is later than second_month.
     
    '''
    month1 = adjust_month(first_month, 0, strft_fmt=None)
    month2 = adjust_month(second_month, 0, strft_fmt=None)
    
    if month1 < month2:
        return -1
    elif month1 == month2:
        return 0
    else: 
        return 1

if __name__ == "__main__":
    
    # test ensure_date_first
    if False:
        dates = [
            '2020-05-05',
            '2020-Jun-12',
            '2002-Jan-31',
            np.nan
            ]
        date_format = 'ymd'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'ymd')
        # ensure_correct_date_format(date_series, 'mdy') # error
        # ensure_correct_date_format(date_series, 'ymd') # error
        pd.to_datetime(date_series)
        pd.to_datetime(date_series, yearfirst = True)
                
        dates = [
            '2020-05-05',
            '2020-06-12',
            '2002-1-31',
            np.nan
            ]
        date_format = 'ymd'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'ymd')
        # ensure_correct_date_format(date_series, 'mdy') # error
        # ensure_correct_date_format(date_series, 'ymd') # error
        pd.to_datetime(date_series)
        pd.to_datetime(date_series, yearfirst = True)
        
        dates = [
            '1/12/2020T23:59:59',
            '2/2/2020 5:9:0',
            '12/3/2020'
            ]
        date_format = 'dmy'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'dmy')
        pd.to_datetime(date_series) # assume mdy
        pd.to_datetime(date_series, dayfirst= True)
        dates = [
            '20-12-2020',
            '12-12-2020',
            '11-14-2020',
            np.nan
            ]
        date_format = 'ymd'
        date_series = pd.Series(dates)
        # ensure_correct_date_format(date_series, 'dmy') # error
        # ensure_correct_date_format(date_series, 'mdy') # error
        
        dates = [
            '1/dec/2020',
            '2/jan/2020',
            '15/mar/2020'
            ]
        date_format = 'dmy'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'dmy')
        pd.to_datetime(date_series)
        pd.to_datetime(date_series, dayfirst= True)
        dates = [
            '1/december/2020',
            '2/january/2020',
            '15/march/2020'
            ]
        date_format = 'dmy'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'dmy')
        pd.to_datetime(date_series)
        pd.to_datetime(date_series, dayfirst= True)
        dates = [
            '30-Oct-20',
            '2-Jan-20',
            ]
        date_format = 'mdy'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'mdy')
        pd.to_datetime(date_series)
        pd.to_datetime(date_series, dayfirst= True)
        
        dates = [
            '1/dec/2020 23:59:59',
            '2/jan/2020',
            '15/mar/2020'
            ]
        date_format = 'dmy'
        date_series = pd.Series(dates)
        ensure_correct_date_format(date_series, 'dmy')
        pd.to_datetime(date_series)
        
        
