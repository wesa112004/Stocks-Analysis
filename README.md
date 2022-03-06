#Stock Analysis With VBA

##Purpose:
  This project looked at 2017 and 2018 stock data and analyze returns using VBA code to loop through the data for Steve. A refactored code was then introduced to see whether the time to execute and run the code was more efficient. Then, an analysis was done to discuss the pros and cons for code refactoring

##Results
-The stock performance in 2017 was superior to 2018 based on the analysis. Specifically, there was special attention given to "DQ" as it earned a return in 2017 of 199.4% vs a negative return in 2018 of - 62.6%. The refactored code ran at roughly .24 seconds and .26 seconds for 2018 and 2017 respectively using the refactored codes. The original codes yeilded .25 seconds and .26 seconds for 2018 and 2017 respectively

<img width="283" alt="VBA_Challenge 2018 Stock Data" src="https://user-images.githubusercontent.com/100330488/156904884-829c614e-a692-4a2e-90a3-043c485f0009.png">
<img width="284" alt="VBA_Challenge 2017 Stock Data" src="https://user-images.githubusercontent.com/100330488/156904892-0c978485-fb26-4990-87b0-81f5035537e0.png">
<img width="760" alt="VBA_Challenge_Code" src="https://user-images.githubusercontent.com/100330488/156904913-a18db42a-6c7a-4252-9e63-7ad0cabb9420.png">

<img width="1512" alt="Screen Shot 2017_refactored 2022-03-05 at 3 22 20 PM" src="https://user-images.githubusercontent.com/100330488/156905010-a0df31a4-df32-4843-8061-6db09a05bd8d.png">

<img width="1512" alt="Screen Shot 2018_refactored 2022-03-05 at 2 03 30 PM" src="https://user-images.githubusercontent.com/100330488/156905014-356c1841-cff4-4fbb-9aa1-a498c9b405ef.png">
<img width="1026" alt="VBA_Challenge_Times 2017 Original code" src="https://user-images.githubusercontent.com/100330488/156905216-a10e4585-7ccb-4537-89f9-4e90c21b4e57.png">
<img width="1283" alt="VBA_Challenge_Time 2018 original" src="https://user-images.githubusercontent.com/100330488/156905236-4411d84b-4a0d-4a5e-b223-ed6fee3a7870.png">


### Summary 
1 In summary, The analysis provided to Steve showed an overall improvement in the market for 2017 with the aforementioned stock "DQ" showing the best results in 2017. The code refactoring process improves the design of the code, making it easier to understand for the developer and ultimately other users. It helps execute programs faster and finds issues within code by cleaning up the original.
  The disadvantages to refactoring can be that it may affect the outcomes of the program. That is, making changes to code can affect the functioning of the program, making changes to the output that may have been unwanmted. In addition, refactoring might be more risky when the application youre working with is large, or if the origianl code is initially already difficult to understand.

2. The pros and Cons apply to refactoring the original VBA script. Refactoring the initial code later down the road could be hard to understand if comments werent made in the original to understand what is happening, therefore its important to specify what each line of code is doing. These were some challenges in this assignment. The pros of refactoring in this case were that it allows one to understand alternative ways to condensing the same information with the same results, and teaches us to refine our approach to analyzing data. 
