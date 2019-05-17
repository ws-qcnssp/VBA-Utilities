# VBALambertWFunction
*Function for calculating the approximation to Lambert W function (product logarithm)*

### Usage
In order to calculate approximation to Lambert W fuction, in the worksheet write function LAMBERTW(x, r, n), which takes three arguments:
 - *x*, which is the main input variable to the function
 - *r*, which is related to approximation of LAMPHI function (so indirectly, also the value of W(x))
 - *n*, which is related to approximation of LAMBERTW

Example input to Excel cell:
```
=LAMBERTW(2000;80;5)
```
Should result in value ca. 5.83673149492073.

Detailed description for this approximation can be found [here](https://mathoverflow.net/questions/57819/best-approximation-to-the-lambertwx-or-explambertwx).

### Limitations
 - Using this function one can calculate approximation of Lambert W function only for real numbers.
 - This code is limited to *r* being integer with max value 80. For greater values Excel throws ARG error.
