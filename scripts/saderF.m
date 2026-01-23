function [F, z, Delta_f]=saderF(z, Delta_f, f_0, k, A)
%SADERF Force deconvolution with the Sader-Jarvis method [1,2] 
%   [F, z, Delta_f] = SADERF(z, Delta_f, f_0, k, A) deconvolutes the force 
%   from a Delta_f frequency shift  vs z distance dataset.
%   Input values
%   z       : distance values in m (vector with length n)
%             z(i) is closer to the surfcace than z(i+1)
%             and z(i) < z(i+1)
%   Delta_f : frequency shift in Hz (vector with length n)
%   f_0     : resonance frequency in Hz
%   k       : spring constant of the cantilever in N/m 
%   A       : amplitude in m
%
%   Output values:
%   F          : deconvoluted force in N (vector with length n-3) 
%   z          : distance values in m (vector with length n-3)
%   Delta_f    : frequency shift in Hz (vector with length n-3)
%
%   based on
%   [1] J. E. Sader and S. P. Jarvis
%       "Accurate formulas for interaction force and energy 
%       in frequency modulation force spectroscopy"
%       Applied Physics Letters, 84, 1801-1803 (2004)  
%   [2] J. E. Sader and S. P. Jarvis
%       Mathematica® notebook for implementation of formulas 
%       http://www.ampc.ms.unimelb.edu.au/afm/bibliography.html#FMAFM. 
%
%   Copyright 2011 Joachim Welker, Esther Illek, Franz J. Giessibl 
%
%   This is an script under the terms of the Creative Commons Attribution 
%   License (http://creativecommons.org/licenses/by/2.0), which permits 
%   unrestricted use, distribution, and reproduction in any medium, 
%   provided the original work is properly cited.
%

% reduced frequency shift Omega
Omega=Delta_f/f_0; 

% derivative of the reduced frequency shift dOmega_dz
dOmega_dz=diff(Omega)./diff(z); 

% adjust length to length of the derviation dOmega_dz
z=z(1:end-1);
Delta_f=Delta_f(1:end-1); 
Omega=Omega(1:end-1); 

for j=1:(numel(z)-2) 
    % start at j+1 due to pole at t=z
    t=z(j+1:end);  
    
    % adjust length of Omega und dOmega_dz to length of t
    Omega_tmp=Omega(j+1:end);
    dOmega_dz_tmp=dOmega_dz(j+1:end);     
    
    % calculate integral Eq.(9) in [1]
    integral=trapz(t, (1+sqrt(A)./(8*sqrt(pi*(t-z(j))))).*Omega_tmp-A^(3/2)./sqrt(2*(t-z(j))).*dOmega_dz_tmp);
                
    % correction terms for t=z from [2]
    corr1 = Omega(j)*(z(j+1)-z(j));                                   
    corr2 = 2*(sqrt(A)/(8*sqrt(pi))) * Omega(j) * sqrt(z(j+1)-z(j)); 
    corr3 = (-2)*(sqrt(A)^3/sqrt(2)) * dOmega_dz(j) * sqrt(z(j+1)-z(j));
    F(j)=2*k*(corr1+corr2+corr3+integral); 
end

% adjust length to length of the force F
z=z(1:numel(F));

% rearrange deconvoluted force to fit input values
F=F';





