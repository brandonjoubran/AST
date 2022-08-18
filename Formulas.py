class Formulas:

    @staticmethod
    def gap_up_perc_open_formula(open_price, prev_close):
        try:
            open_price_int = float(open_price)
            prev_close_int = float(prev_close)
            if(prev_close_int == 0):
                return 0
            return ((open_price_int - prev_close_int) / prev_close_int) * 100
        except:
            print("Error occured in gap_up_perc_open_formula(). Values given were open_price={} and prev_close={}".format(open_price, prev_close))
        return 'N/A'


    @staticmethod
    def gap_up_perc_premarket_formula(prev_close, premarket_high):
        #premarket_high = self.split_premarket_high(premarket_high)
        try:
            premarket_high_float = float(premarket_high)
            prev_close_float = float(prev_close)
            if (prev_close_float == 0):
                return 0
            return ((premarket_high_float - prev_close_float)/prev_close_float)*100
        except:
            print("Error occured in gap_up_perc_premarket_formula(). Values given were prev_close={} and premarket_high={}".format(prev_close, premarket_high))
        return 'N/A'

    @staticmethod
    def gap_perc_maintained_by_open(gap_up_perc_open, gap_up_perc_premarket):
        try:
            if(gap_up_perc_premarket == 0):
                return 0
            result = (gap_up_perc_open/gap_up_perc_premarket)*100
            return result
        except:
            print("Error occured in gap_perc_maintained_by_open(). Values given were gap_up_perc_open={} and gap_up_perc_premarket={}".format(
                gap_up_perc_open, gap_up_perc_premarket
            ))
        return 'N/A'

    @staticmethod
    def spike_perc(days_high, open_price):
        try:
            open_price_float = float(open_price)
            days_high_float = float(days_high)
            if(open_price_float == 0):
                return 0
            return ((days_high_float - open_price_float) / open_price_float) * 100
        except:
            print("Error occured in spike_perc(). Values given were days_high={} and open_price={}".format(days_high, open_price))
        return 'N/A'

    @staticmethod
    def fail_perc(days_low, open_price):
        try:
            open_price_float = float(open_price)
            days_low_float = float(days_low)
            if(open_price_float == 0):
                return 0
            return ((open_price_float - days_low_float) / open_price_float) * 100
        except:
            print("Error occured in fail_perc(). Values given were days_low={} amd open_price={}".format(days_low, open_price))
        return 'N/A'

    @staticmethod
    def perc_of_float_trade(float, vol):
        try:
            if(float == 0):
                return 0
            return (vol/float)*100
        except:
            print("Error occured in perc_of_float_trade(). Values given were float={} amd vol={}".format(float, vol))
        return 'N/A'

    @staticmethod
    def pullback_from_pm_high_to_open(premarket_high, open_price):
        try:
            open_price_float = float(open_price)
            premarket_high_float = float(premarket_high)
            if(premarket_high_float == 0):
                return 0
            return ((premarket_high_float - open_price_float)/premarket_high_float)*100
        except:
            print("Error occured in pullback_from_pm_high_to_open(). Values given were premarket_high={} amd open_price={}".format(premarket_high, open_price))
        return 'N/A'