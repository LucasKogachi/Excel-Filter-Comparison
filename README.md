# Excel Filter & Comparison

## Overview
This Excel Project is a Simple Tool using Excel VBA to Filter and Compare DATA.

## Why Use this Tool?

This tool is useful when you work in a company and the machine only has access to the company system and the Microsoft Office. Sometimes we stumble trying to do some simple things with the DATA and we only have Excel to help us, but what we need is a little more complicated than the usual, but still simple enough.

### Example

A Company needs to change many programs to accept the new changes in the system, but each program can call other subprograms and sometimes it’s necessary to change those subprograms too. Programs can also have subprograms in common.

Every program has a list of subprograms used. To keep record of changes, there is a list of subprograms already changed and working, there is also another list with subprograms with new changes being evaluated and finally a list with subprograms of a program your team is working on.

Using the tool, you copy the list of subprograms of a program your team is working on to the Input Sheet. To filter that you copy the list of subprograms already changed and working to a new Filter Sheet, in another new filter sheet you copy the list with subprograms with new changes being evaluated. Finally you can add a new Comparison Sheet with a list of who of your team is working with what subprogram.

Running the tool, Output Sheet shows what subprograms your team needs to focus on and who of your team is working with what subprogram.

This is just an example, you can find new uses for this tool.

## How it Works

There are 3 Basic Sheets:
+ Menu
+ Input
+ Output

To understand how it works, it’s better to look into what the project can do.

### Menu Sheet

Here you can:
+ Add Filter Sheets
+ Add Comparison Sheets
+ Run those Filters and Comparisons
+ Remove Duplicates from Input Sheet
+ Sort Input Sheet(Ascending)
+ Sort Output Sheet(Ascending)

The difference between Filter and Comparison is that Filter Paints Input and the rest goes to Output (Rows not Painted), while Comparison just Paints Output.

### Filters Sheets

You can choose the number of columns to be compared with the Input Sheet and choose a color that will fill the Input Cells when equal, the rows that were NOT a match will be copied to the Output Sheet.

### Comparisons Sheets

You can choose the number of columns to be compared with the Output Sheet and choose a color that will fill the Output Cells when equal.

### Examples

There is a pdf for examples and more details

## Getting Started

1. Download the .xlsm file

2. Open it

3. Enable Macros (File -> Options -> Trust Center -> Trust Center Settings... -> Trusted Locations -> Add File Location -> Re-open File)

4. Enjoy it
