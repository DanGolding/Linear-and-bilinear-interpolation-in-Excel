# Linear-and-bilinear-interpolation-in-Excel
VBA code of worksheet functions for linear and bilinear interpolation based on the signature of [`interp1`](http://uk.mathworks.com/help/matlab/ref/interp1.html) and [`interp2`](http://uk.mathworks.com/help/matlab/ref/interp2.html) in MATLAB.

The code is in [Interpolation.bas](https://github.com/DanGolding/Linear-and-bilinear-interpolation-in-Excel/blob/master/Interpolation.bas) and an [example spreadheet](https://github.com/DanGolding/Linear-and-bilinear-interpolation-in-Excel/blob/master/Interpolation%20example.xlsm) is also provided.

### Linear Interpolation
Example of using the `interp1` function:

<p align="center">
  <img src="/Images/linear_interpolation_example.png" />
</p>

### Bilinear Interpolation
Example of using the `interp2` function:

<p align="center">
  <img src="/Images/bilinear_interpolation_example.png" />
</p>

note for `inrerp2`, the first arguemnt, `xAxis`, must be a _vertical_ range, and the second argument, `yAxis`, must be a _horizontal_ range.
