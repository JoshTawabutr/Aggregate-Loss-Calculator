# Aggregate Loss Calculator with VBA

This repository contains Visual Basic code that computes the discrete aggregate losses distribution from the inputs of discretized frequency and severity distributions. The user is recommended to run the program on Microsoft Excel. 

___
## Overview

- [Project Description](#Project-Description)
- [Inputs](#Inputs)
- [Computation and Output](#Computation-and-Output)
- [Underlying Mathematics](#Underlying-Mathematics)
- [Authorship](#Authorship)

___
## Project Description

A vital part to sustaining an insurance company consists of two estimations: (i) the amount required to reserve for future claim payments, and (ii) the proper prices for future policies. Both of them are performed based on past claim data, which have been shown to be modeled most accurately by separating them into frequency (number of claims per exposure) and severity (amount per claim), each modeled by a probability distribution. Afterwards, the frequency and severity distributions are combined into the aggregate losses (total losses per exposure) distribution.

This is where the project comes in. This project provides a Visual Basic code that operates on Microsoft Excel and takes as inputs the discrete (or discretized) frequency and severity distributions. The input must be in the format of a table of discrete frequency/severity values, each provided with its probability function. The user can provide these inputs through the included input forms, external scripts or manual inputs. The program then computes the resulting aggregate losses distribution and displays it in the table format on a separate Excel worksheet. The details of how each step works is provided below. 


[Back to Overview](#overview)
___
## Inputs

The Excel file contains three worksheets: Frequency, Severity and Aggregate Losses. As the names suggest, each sheet holds the table representing the discrete probability function of the corresponding quantity. For instance, the original Frequency worksheet is displayed below.
<p align="center">
<img src = "Freq_BlankSheet.png" width = "500"></img>
</p>
Besides the table mentioned above, the worksheet has two buttons. The top button launches a form that allows the user to add rows to the table, while the bottom button launches a form that allows the user to execute aggregate losses computation. The Aggregate Losses worksheet only has the latter button. 

<br> Once the user clicks on the "Input frequency probability function" button, the input form shown below will be launched. 
<p align="center">
<img src = "Freq_InputForm.png" width = "500"></img>
</p>
The form allows the user to input the frequency value and its corresponding probability, one pair at a time. The form also controls for the proper value of probability input. For example, if a user inputs a negative probability by mistake, the following warning box will be displayed. 
\
<p align="center">
<img src = "Freq_InputFormFilled1.png" width = "450"></img>      
<img src = "Freq_InputWarning1.png" width = "450"></img>
</p>
Furthermore, if the input probability would make the total probability sum to a number above one, the program will display the following warning box. 
\
<p align="center">
<img src = "Freq_InputFormFilled2.png" width = "450"></img>      
<img src = "Freq_InputWarning2.png" width = "450"></img>
</p>
Once the whole frequency distribution is given, the user proceeds to do the same with the severity distribution. This is done in the Severity worksheet shown below, together with the severity input form. 
\
<p align="center">
<img src = "Sev_BlankSheet.png" width = "450"></img>      
<img src = "Sev_InputForm.png" width = "450"></img>
</p>
Alternative ways the user can input the frequency and severity distributions include the use of an external script to fill the tables in the corresponding worksheets. The user can also put the numbers in manually. However, if the inputs are given in these alternative ways, the program will not check for errors in the input until the actual computation is launched.


[Back to Overview](#overview)
___
## Computation and Output

Blah


[Back to Overview](#overview)
___
## Underlying Mathematics

Blah


[Back to Overview](#overview)
___
## Authorship

This project is developed by Yossathorn (Josh) Tawabutr. For more information, please contact: tawabutr.1@osu.edu





Personal project <br>
Overview

Motivation -- importance of aggregate loss model

Input <br>
Can click and use the form (show frequency sheet and its input form when launched) <br>
Check for positive input (show screaming box) <br>
Check that prob sum does not exceed 1 (show screaming box) <br>
Alternatively, can copy and paste from external worksheet (show pasted numbers) <br>
Same for severity dist (show severity sheet) <br>

Computation <br>
Can click the form on aggregate losses sheet (Show the sheet and then the launched form) <br>
Can choose from 3 output formats: explain <br>
(Show overwriting warning box) <br>
The program will check if the given frequency and severity prob fns each sums to one. (Show such the warning box) <br>
(Show success box) <br>
(Show finished sheet)

Some brief conclusion

Remark about project creator
