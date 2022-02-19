# Suwarin

Version 1.0.0

![demo](/images/suwarin.gif)

[Japanese README](README.ja.md)

[GitHub](https://github.com/taidalog/Suwarin)

## Table of Contents

1. [Introduction](#Introduction)
1. [Features](#Features)
1. [Installation](#Installation)
1. [Usage](#Usage)
1. [Notes](#Notes)
1. [License](#License)


## Introduction

Excel macro to make a seating chart. **This is not for changing the seating arrangement, or changing the seating order, but for making a seating chart with a specified order of participants.** It will be helpful when you are making several sheets of onetime seating charts.


## Features

- Making a seating chart only by copy and paste participants and clicking twice.
- Using this macro from the context menu.
- Specifying seats not to be used.
- Designing the chart as you like to some extent.

Copy and paste participants to the column next to the seating chart, then click button in the context menu. Then the participants will be input to the chart.

You can specify the seats to skip. Enter lowercase "x" to the upper left cell of the seats to skip.

The range of the seating chart range and each seat are specified by the ruled line instead of cell address. So if you draw lines according to the rules, you can design the chart as you like to some extent.


## Installation

1. Download `main.bas`
1. Import the module, or copy & paste the second line and below.
1. To Thisworkbook module, add the code below:  
    ```
    Private Sub Workbook_Open()
        Call AddToContextMenu
    End Sub
    ```
1. Save the file as `.xlsm` (any name is OK)
1. Reopen the file


## Usage

1. Draw the ruled lines to make a seating chart
1. Enter participants into the cells two cells away from the top right cell of the seating chart
![layout](/images/suwarin01.en.png)
1. Right click on a cell (anywhere) and click a buttom in the menu


## Notes

- The range of the seating chart range and each seat are specified by the ruled line instead of cell address. Make sure that the lines are within the rules below:
    - Draw a ruled line on the outer frame of the seating chart without interruption.
    ![About ruled lines](/images/suwarin02.en.png)
    - Do not draw extra ruled lines beyond the range of the seating chart.
    ![About ruled lines](/images/suwarin03.en.png)
    - Separate each seat with a ruled line without interruption.
    ![About ruled lines](/images/suwarin04.en.png)
    - Do not draw a ruled line inside each seat.
    - Make the number of rows and columns of each seat the same.
    ![About ruled lines](/images/suwarin05.en.png)
- The range to enter the participants into is designated. The top of the range is two cells away from the top right cell of the seating chart. If the chart range is `B3:M16`, the top right cell is `L3`, so enter the participants into `O3`, `O4`, `O5` ....  
![About cells to enter participants](/images/suwarin01.en.png)
- Each participant will be entered into the top left cell of each seat. Nothing will be done to the other cells. You can enter functions.
- Make sure that only one seating chart exists on a worksheet.


## License

Copyright 2022 taidalog

Suwarin is licensed under the MIT License.
