function [F, z, Delta_f]=matrixF(z, Delta_f, f_0, k, A)
%MATRIXF Force deconvolution with the matrix method [1] 
%   [F, z, Delta_f]=MATRIXF(z, Delta_f, f_0, k, A) deconvolutes the force 
%   from a df frequency shift vs z distance dataset.
%   Input values
%   z       : distance values in m (vector with length n)
%             z(i) is closer to the surfcace than z(i+1)
%             and z(i) < z(i+1).
%             z values must be equidistant.
%   Delta_f : frequency shift in Hz (vector with length n)
%   f_0     : resonance frequency in Hz
%   k       : spring constant of the cantilever in N/m 
%   A       : amplitude in m
%
%   Output values
%   F          : deconvoluted force in N (vector with length n) 
%   z          : distance values in m (vector with length n)
%   Delta_f    : frequency shift in Hz (vector with length n)
%   
%   based on
%   [1] F. J. Giessibl
%       "A Direct Method to Calculate Tip-Sample Forces from
%       Frequency Shifts in Frequency-Modulation Atomic Force Microscopy"
%       Applied Physics Letters 78, 123-125 (2001)
%
%   Copyright 2011 Joachim Welker, Esther Illek, Franz J. Giessibl 
%
%   This is a script under the terms of the Creative Commons Attribution 
%   License (http://creativecommons.org/licenses/by/2.0), which permits 
%   unrestricted use, distribution, and reproduction in any medium, 
%   provided the original work is properly cited.
%

% rearrange Delta_f values to fit the algorithm [1]
Delta_f=Delta_f(end:-1:1); 

% step width of the data
Delta=z(2)-z(1); 

% integer expression of the amplitude in terms of the step width Delta
alpha=round(A/Delta);

% initalize matrix W
N=numel(z); 
W=zeros(N,N);

% calculate elements of matrix W, Eq.(6) in [1]
for i=1:N
    x=max(i-2*alpha,1);
    for j=x:i
        W(i,j)=(f_0/2/k)*(2/pi/A)*2/(2*alpha+1)*(sqrt((2*alpha+1)*(i-j+1)-(i-j+1)^2)-sqrt((2*alpha+1)*(i-j)-(i-j)^2));
    end
end

% calculate Eq.(7) in [1]
F=W\Delta_f;

% rearrange deconvoluted force to fit input values
Delta_f=Delta_f';
F=F(end:-1:1);
