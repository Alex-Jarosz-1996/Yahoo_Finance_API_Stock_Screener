import yfinance as yf

class Stock:
    def __init__(self, ticker): 

        num_dp = 3
        
        # setting up the below metrics
        self.obj  = yf.Ticker(f'{ticker}')
        self.info = self.obj.info
        
        # stock price metrics
        self.ticker          = ticker
        self.price           = float(self.info['currentPrice']) if self.info['currentPrice'] is not None else None
        self.marketCap       = float(self.info['marketCap']) if self.info['marketCap'] is not None else None
        self.numShares       = float(self.info['sharesOutstanding']) if self.info['sharesOutstanding'] is not None else None
        self.yearlyLowPrice  = float(self.info['fiftyTwoWeekLow']) if self.info['fiftyTwoWeekLow'] is not None else None
        self.yearlyHighPrice = float(self.info['fiftyTwoWeekHigh']) if self.info['fiftyTwoWeekHigh'] is not None else None
        self.fiftyDayMA      = float(self.info['fiftyDayAverage']) if self.info['fiftyDayAverage'] is not None else None
        self.twoHundredDayMA = float(self.info['twoHundredDayAverage']) if self.info['twoHundredDayAverage'] is not None else None

        # value metrics
        self.acquirersMultiple     = round(float(self.info['enterpriseValue'] / self.info['netIncomeToCommon']), num_dp) if (self.info['enterpriseValue'] / self.info['netIncomeToCommon']) is not None else None
        self.currentRatio          = round(float(self.info['currentRatio']), num_dp) if self.info['currentRatio'] is not None else None
        self.enterpriseValue       = round(float(self.info['enterpriseValue']), num_dp) if self.info['enterpriseValue'] is not None else None
        self.eps                   = round(float(self.info['trailingEps']), num_dp) if self.info['trailingEps'] is not None else None
        self.evToEBITDA            = round(float(self.info['enterpriseToEbitda']), num_dp) if self.info['enterpriseToEbitda'] is not None else None
        self.evToOperatingCashFlow = round(float(self.info['enterpriseValue'] / self.info['operatingCashflow']), num_dp) if (self.info['enterpriseValue'] / self.info['operatingCashflow']) is not None else None
        self.evToRev               = round(float(self.info['enterpriseToRevenue']), num_dp) if self.info['enterpriseToRevenue'] is not None else None
        self.peRatioTrail          = round(float(self.info['trailingPE']), num_dp) if self.info['trailingPE'] is not None else None
        self.peRatioForward        = round(float(self.info['forwardPE']), num_dp) if self.info['forwardPE'] is not None else None
        self.priceToSales          = round(float(self.info['priceToSalesTrailing12Months']), num_dp) if self.info['priceToSalesTrailing12Months'] is not None else None
        self.priceToBook           = round(float(self.info['priceToBook']), num_dp) if self.info['priceToBook'] is not None else None

        # dividend metrics
        self.dividendYield = float(self.info['trailingAnnualDividendYield']) if self.info['trailingAnnualDividendYield'] is not None else None
        self.dividendRate  = float(self.info['dividendRate']) if self.info['dividendRate'] is not None else None
        self.exDivDate     = str(self.info['exDividendDate']) if self.info['exDividendDate'] is not None else None
        self.lastDivVal    = float(self.info['lastDividendValue']) if self.info['lastDividendValue'] is not None else None
        self.payoutRatio   = str(self.info['payoutRatio']) if self.info['payoutRatio'] is not None else None

        # # balance sheet metrics
        self.bookValue         = round(float(self.info['bookValue']), num_dp) if self.info['bookValue'] is not None else None
        self.bookValPerShare   = round(float(self.info['bookValue'] / self.info['sharesOutstanding']), num_dp) if (self.info['bookValue'] / self.info['sharesOutstanding']) is not None else None
        self.cash              = round(float(self.info['totalCash']), num_dp) if self.info['totalCash'] is not None else None
        self.cashPerShare      = round(float(self.info['totalCashPerShare']), num_dp) if self.info['totalCashPerShare'] is not None else None
        self.cashToMarketCap   = round(float(self.info['totalCash'] / self.info['marketCap']), num_dp) if (self.info['totalCash'] / self.info['marketCap']) is not None else None
        self.cashToDebt        = round(float(self.info['totalCash'] / self.info['totalDebt']), num_dp) if (self.info['totalCash'] / self.info['totalDebt']) is not None else None
        self.debt              = round(float(self.info['totalDebt']), num_dp) if self.info['totalDebt'] is not None else None
        self.debtToMarketCap   = round(float(self.info['totalDebt'] / self.info['marketCap']), num_dp) if (self.info['totalDebt'] / self.info['marketCap']) is not None else None
        self.debtToEquityRatio = round(float(self.info['debtToEquity']), num_dp) if self.info['debtToEquity'] is not None else None
        self.quickRatio        = round(float(self.info['quickRatio']), num_dp) if self.info['quickRatio'] is not None else None
        self.returnOnAssets    = round(float(self.info['returnOnAssets']), num_dp) if self.info['returnOnAssets'] is not None else None
        self.returnOnEquity    = round(float(self.info['returnOnEquity']), num_dp) if self.info['returnOnEquity'] is not None else None
        self.totalAssets       = round(float(self.info['totalAssets']), num_dp) if self.info['totalAssets'] is not None else None

        # income related
        self.ebitda              = round(float(self.info['ebitda']), num_dp) if self.info['ebitda'] is not None else None
        self.ebitdaMargins       = round(float(self.info['ebitdaMargins']), num_dp) if self.info['ebitdaMargins'] is not None else None
        self.ebitdaPerShare      = round(float(self.info['ebitda']), num_dp) if self.info['ebitda'] is not None else None
        self.earningsGrowth      = round(float(self.info['earningsGrowth']), num_dp) if self.info['earningsGrowth'] is not None else None
        self.grossMargins        = round(float(self.info['grossMargins']), num_dp) if self.info['grossMargins'] is not None else None
        self.grossProfit         = round(float(self.info['grossProfits']), num_dp) if self.info['grossProfits'] is not None else None
        self.grossProfitPerShare = round(float(self.info['grossProfits'] / self.info['sharesOutstanding']), num_dp) if (self.info['grossProfits'] / self.info['sharesOutstanding']) is not None else None
        self.netIncome           = round(float(self.info['netIncomeToCommon']), num_dp) if self.info['netIncomeToCommon'] is not None else None
        self.netIncomePerShare   = round(float(self.info['netIncomeToCommon'] / self.info['sharesOutstanding']), num_dp) if (self.info['netIncomeToCommon'] / self.info['sharesOutstanding']) is not None else None
        self.operatingMargin     = round(float(self.info['operatingMargins']), num_dp) if self.info['operatingMargins'] is not None else None
        self.profitMargin        = round(float(self.info['profitMargins']), num_dp) if self.info['profitMargins'] is not None else None
        self.revenue             = round(float(self.info['totalRevenue']), num_dp) if self.info['totalRevenue'] is not None else None
        self.revenueGrowth       = round(float(self.info['revenueGrowth']), num_dp) if self.info['revenueGrowth'] is not None else None
        self.revenuePerShare     = round(float(self.info['revenuePerShare'] / self.info['sharesOutstanding']), num_dp) if (self.info['revenuePerShare'] / self.info['sharesOutstanding']) is not None else None

        # cash flow related
        self.fcf               = round(float(self.info['freeCashflow']), num_dp) if self.info['freeCashflow'] is not None else None
        self.fcfToMarketCap    = round(float(self.info['freeCashflow'] / self.info['marketCap']), num_dp) if (self.info['freeCashflow'] / self.info['marketCap']) is not None else None
        self.fcfPerShare       = round(float(self.info['freeCashflow'] / self.info['sharesOutstanding']), num_dp) if (self.info['freeCashflow'] / self.info['sharesOutstanding']) is not None else None
        self.ocf               = round(float(self.info['operatingCashflow']), num_dp) if self.info['operatingCashflow'] is not None else None
        self.ocfToRevenueRatio = round(float(self.info['operatingCashflow'] / self.info['totalRevenue']), num_dp) if (self.info['operatingCashflow'] / self.info['totalRevenue']) is not None else None
        self.ocfToMarketCap    = round(float(self.info['operatingCashflow'] / self.info['marketCap']), num_dp) if (self.info['operatingCashflow'] / self.info['marketCap']) is not None else None
        self.ocfPerShare       = round(float(self.info['operatingCashflow'] / self.info['sharesOutstanding']), num_dp) if (self.info['operatingCashflow'] / self.info['sharesOutstanding']) is not None else None
        self.fcfToEV           = round(float(self.info['freeCashflow'] / self.info['enterpriseValue']), num_dp) if (self.info['freeCashflow'] / self.info['enterpriseValue']) is not None else None
        self.ocfToEV           = round(float(self.info['operatingCashflow'] / self.info['enterpriseValue']), num_dp) if (self.info['operatingCashflow'] / self.info['enterpriseValue']) is not None else None

    def __repr__(self):
        return self

    def __str__(self):
        return self 

    def get_all_keys(self):
        """
        returns all possible keys that is available to the developer
        """
        for i in self.info.keys():
            print(i)

    def print_all_class_properties(self):
        """
        prints all properties
        """
        print(', '.join("%s: %s" % item for item in vars(Stock(f'{self.ticker}')).items()))


