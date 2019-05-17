# Utilities for sets
*A collection of simple functions for sets of data*

### Functions
Currently, following functions are implemented in this module:
 - SUMOFMULT
 - SUMOFMULT2
 - SUMOFDIV2
 - SUMOFMARKED
 - AVERAGEOFMARKED
 
### Usage

#### SUMOFMULT
This function takes multiple identical-sized ranges as an input and calculates sum of multiplications of *i*th row and *j*th column cell of each range:
![           N          N                             
        __  rm row __  rm col                       
S  =   \          \           a   cdot b   cdot dots
       /__ i = 1  /__ j = 1    ij       ij          
](http://www.sciweavers.org/tex2img.php?eq=S%20%3D%20%20%5Csum_%7Bi%3D1%7D%5E%7BN_%7B%5Crm%20row%7D%7D%5Csum_%7Bj%3D1%7D%5E%7BN_%7B%5Crm%20col%7D%7D%20a_%7Bij%7D%20%5Ccdot%20b_%7Bij%7D%20%5Ccdot%20%5Cdots&bc=White&fc=Black&im=jpg&fs=12&ff=arev&edit=0 "SUMOFMULT Equation")
