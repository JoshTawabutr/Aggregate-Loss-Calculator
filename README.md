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

Blah


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
