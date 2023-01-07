from datetime import datetime

class timectrl():
    def __init__(self):
        self.__starttime = 0
        self.__endtime = 0

    def getTimeStamp():
        today = datetime.today()
        return today.strftime('%Y-%m-%d %H:%M:%S.' + str(today.microsecond))
