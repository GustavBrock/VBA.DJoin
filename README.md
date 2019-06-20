# VBA.DJoin #
![Help](https://raw.githubusercontent.com/GustavBrock/VBA.DJoin/master/images/EE%20Header.png)

### Join (concat) values from one field from a table or query  ###
As you can union records, you can join field values. DJoin offers increased speed and flexibility compared to the ancient ConcactRelated and similar functions. Further, it offers better read-out of Multi-Value fields. 

For decades, functions have been around to solve the simple task of joining (concatenating) the values from a single field of many records to a single field value - as illustrated by the title picture, where the values from two keys (left) are joined into one field for each key (right) with a delimiter (or separator), here a *space*.

DJoin is named such to signal the familiarity with the native *domain aggregate functions* - DLookup, DCount, etc. - as it aggregates the values from one field from many records to one string - much like *Join* does for an array.

Areas for improvement and added flexibility:

* Better speed, indeed when browsing
* A wider choice of source types - like pure SQL
* Caching of results for vastly improved speed for repeated calls

The function SpeedTest (from the attached demo below) reveals - for a large table of ~170,000 records having 424 keys - a speed improvement on the first run of about 25%. Repeated calls run about 16 times faster - a dramatic speed increase:

![Help](https://raw.githubusercontent.com/GustavBrock/VBA.DJoin/master/images/SpeedTest.PNG)

### Code ###
Code has been tested with both 32-bit and 64-bit *Microsoft Access 2016* and *365*.

### Documentation ###
Full documentation can be found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.DJoin/master/images/EE%20Logo.png) 

[Join (concat) values from one field from a table or query](https://www.experts-exchange.com/articles/33612/Join-concat-values-from-one-field-from-a-table-or-query.html)

Included is a Microsoft Access example application.