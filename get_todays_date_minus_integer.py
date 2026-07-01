# The purpose of this is to send an xkcd comic once a day, in creation date order, 
# starting with the last one that I've read

from datetime import date, timedelta


# This site helped get as integer: https://www.dataquest.io/blog/python-datetime-tutorial/
last_comic_read = 340
d = timedelta(days=last_comic_read)    # Gets number of days between last comic read and today                                  
next_comic = d.days    # Creates an integer from the string
print(d)
print(next_comic)                
