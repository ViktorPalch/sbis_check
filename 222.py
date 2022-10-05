import datetime
s='10 days, 0:06:15.854000'
# s='1 day, 0:03:01.995000'

print(len(str(s)), str(s[8]))
print(str(s)[0:2], (str(s))[7:])
d16 = datetime.datetime.strptime(
    (str(s))[0:2] + ' ' + (str(s))[7:],
    '%d %H:%M:%S.%f')
d16 = datetime.timedelta(days=d16.day, hours=d16.hour, minutes=d16.minute, seconds=d16.second,
                         microseconds=d16.microsecond)

print(d16)