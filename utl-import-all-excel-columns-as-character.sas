Import all excel columns as character;                                                                              
                                                                                                                    
github                                                                                                              
https://tinyurl.com/yx6d2z9t                                                                                        
https://github.com/rogerjdeangelis/utl-import-all-excel-columns-as-character                                        
                                                                                                                    
macros                                                                                                              
https://tinyurl.com/y9nfugth                                                                                        
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories                          
                                                                                                                    
  Algorithm                                                                                                         
                                                                                                                    
    1.   Read the column names as data this will force character type for all columns                               
         Valid column names cannot begin with a number                                                              
    2.   Use options header=no and mixed=yes options on libname                                                     
    3.   On import drop second row(has the names)                                                                   
    4.   rename cols f1-f5(libname default names when header=n) to x1-x5(orginal names)                             
         Notes you can rename to arbitrary names using just a few lines of code                                     
         %array and $do_over                                                                                        
                                                                                                                    
   Two solutions                                                                                                    
                                                                                                                    
        a. When we have a named range and column names x1-x5                                                        
        b. Arbitrary: Without a named range and arbitrary column names                                              
*_                   _                                                                                              
(_)_ __  _ __  _   _| |_                                                                                            
| | '_ \| '_ \| | | | __|                                                                                           
| | | | | |_) | |_| | |_                                                                                            
|_|_| |_| .__/ \__,_|\__|                                                                                           
        |_|                                                                                                         
;                                                                                                                   
                                                                                                                    
%utlopts;                                                                                                           
libname xel clear;                                                                                                  
%utlfkil(d:/xls/have.xlsx); * delete if exists;                                                                     
                                                                                                                    
libname xel "d:/xls/have.xlsx";                                                                                     
                                                                                                                    
data xel.have;                                                                                                      
                                                                                                                    
 array TFS[5] X1-X5;                                                                                                
 do rec = 1 to 5;                                                                                                   
                                                                                                                    
    do idx=1 to dim(tfs);                                                                                           
       tfs[idx] = (RAND("Bernoulli", .25)=1);                                                                       
    end;                                                                                                            
    output;                                                                                                         
    drop rec;                                                                                                       
 end;                                                                                                               
                                                                                                                    
 drop idx;                                                                                                          
 stop;                                                                                                              
                                                                                                                    
run;quit;                                                                                                           
                                                                                                                    
libname xel clear;                                                                                                  
                                                                                                                    
Excel d:/xls/have.xlsx                                                                                              
                                                                                                                    
    +--------------+                                                                                                
    |A |B |C |D |D |                                                                                                
    |--+--+--+--+--|                                                                                                
  1 |X1|X2|X3|X4|X5|                                                                                                
    |--+--+--+--+--|                                                                                                
  2 | 0| 0| 0| 0| 0|                                                                                                
    |--+--+--+--+--|                                                                                                
  3 | 1| 0| 1| 0| 0|                                                                                                
    |--+--+--+--+--|                                                                                                
  4 | 0| 0| 1| 0| 1|                                                                                                
    |--+--+--+--+--|                                                                                                
  5 | 0| 0| 0| 0| 0|                                                                                                
    |--+--+--+--+--|                                                                                                
  5 | 0| 1| 1| 0| 0|                                                                                                
    +--------------+                                                                                                
*            _               _                                                                                      
  ___  _   _| |_ _ __  _   _| |_                                                                                    
 / _ \| | | | __| '_ \| | | | __|                                                                                   
| (_) | |_| | |_| |_) | |_| | |_                                                                                    
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                   
                |_|                                                                                                 
;                                                                                                                   
                                                                                                                    
Variables in Creation Order                                                                                         
                                                                                                                    
#    Variable    Type    Len                                                                                        
                                                                                                                    
1    X1          Char      2                                                                                        
2    x2          char      2                                                                                        
2    X2          Char      2                                                                                        
3    X3          Char      2                                                                                        
4    X4          Char      2                                                                                        
5    X5          Char      2                                                                                        
                                                                                                                    
                                                                                                                    
WORK.WANT total obs=5                                                                                               
                                                                                                                    
  X1    X2    X3    X4    X5                                                                                        
                                                                                                                    
  0     0     1     1     0                                                                                         
  0     0     0     0     0                                                                                         
  0     0     0     0     0                                                                                         
  0     0     0     0     0                                                                                         
  0     1     0     1     0                                                                                         
                                                                                                                    
*                                                                                                                   
 _ __  _ __ ___   ___ ___  ___ ___                                                                                  
| '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                 
| |_) | | | (_) | (_|  __/\__ \__ \                                                                                 
| .__/|_|  \___/ \___\___||___/___/                                                                                 
|_|                                                                                                                 
;                                                                                                                   
*                                         _                                                                         
  __ _     _ __   __ _ _ __ ___   ___  __| |  _ __ __ _ _ __   __ _  ___                                            
 / _` |   | '_ \ / _` | '_ ` _ \ / _ \/ _` | | '__/ _` | '_ \ / _` |/ _ \                                           
| (_| |_  | | | | (_| | | | | | |  __/ (_| | | | | (_| | | | | (_| |  __/                                           
 \__,_(_) |_| |_|\__,_|_| |_| |_|\___|\__,_| |_|  \__,_|_| |_|\__, |\___|                                           
                                                              |___/                                                 
;                                                                                                                   
                                                                                                                    
proc datasets lib=work nolist;                                                                                      
  delete want;                                                                                                      
run;quit;                                                                                                           
                                                                                                                    
libname xel "d:/xls/have.xlsx" header=no mixed=yes;                                                                 
                                                                                                                    
data want (rename=(f1-f5=x1-x5));                                                                                   
                                                                                                                    
  set xel.have;                                                                                                     
  if (_n_ ne 1) ;                                                                                                   
                                                                                                                    
run;quit;                                                                                                           
                                                                                                                    
libname xel clear;                                                                                                  
                                                                                                                    
*_                   _     _ _                                                                                      
| |__      __ _ _ __| |__ (_) |_ _ __ __ _ _ __ _   _                                                               
| '_ \    / _` | '__| '_ \| | __| '__/ _` | '__| | | |                                                              
| |_) |  | (_| | |  | |_) | | |_| | | (_| | |  | |_| |                                                              
|_.__(_)  \__,_|_|  |_.__/|_|\__|_|  \__,_|_|   \__, |                                                              
                                                |___/                                                               
;                                                                                                                   
                                                                                                                    
* delete all macro arrays if they exist from prior run;                                                             
%utlnopts;                                                                                                          
%symdel rer / nowarn;                                                                                               
%deleteMacArray(f_names,1);                                                                                         
%deleteMacArray(source_names,1);                                                                                    
                                                                                                                    
* delete imported sas table if it exists;                                                                           
proc datasets lib=work nolist;                                                                                      
  delete want;                                                                                                      
run;quit;                                                                                                           
                                                                                                                    
libname xel "d:/xls/have.xlsx";                                                                                     
                                                                                                                    
%array(source_names,values=%varlist(xel.have));                                                                     
                                                                                                                    
%array(f_names,values=f1-f&source_namesn);                                                                          
                                                                                                                    
%let ren = %do_over(f_names source_names,phrase=%Str(?f_names=?source_names ));                                     
                                                                                                                    
%put &ren;                                                                                                          
                                                                                                                    
/*                                                                                                                  
Copy this from log and paste in code or use do_over directly                                                        
                                                                                                                    
f1=X1 f2=X2 f3=X3 f4=X4 f5=X5                                                                                       
*/                                                                                                                  
                                                                                                                    
* DO THIS ;                                                                                                         
                                                                                                                    
libname xel "d:/xls/have.xlsx" header=no mixed=yes;                                                                 
                                                                                                                    
data want ;                                                                                                         
                                                                                                                    
  set xel.'have$'n;                                                                                                 
                                                                                                                    
  if (_n_ ne 1) ;                                                                                                   
                                                                                                                    
  rename f1=X1 f2=X2 f3=X3 f4=X4 f5=X5 ;                                                                            
                                                                                                                    
run;quit;                                                                                                           
                                                                                                                    
* OR THIS;                                                                                                          
                                                                                                                    
data want ;                                                                                                         
                                                                                                                    
  set xel.'have$'n;                                                                                                 
  if (_n_ ne 1) ;                                                                                                   
  rename                                                                                                            
    &ren                                                                                                            
  ;                                                                                                                 
                                                                                                                    
run;quit;                                                                                                           
                                                                                                                    
* clean up macro arrays;                                                                                            
                                                                                                                    
%deleteMacArray(f_names,1);                                                                                         
%deleteMacArray(source_names,1);                                                                                    
                                                                                                                    
