# Head Pipe Loss

Head Pipe Loss is a Python project for calculating the pressure loss in an HVAC PPR piping system.
The total piping system is divided into multiple nodes, based on pipe's diameter change.
Each node's/section's pipe loss is calculated based on the water flow rate and pipe and fittings' characteristics.
Total Loss is calculated for both Supply and Return side.
The characteristics of pipes and fittings are based on the Aquatherm PPR series, but can be modified for other manufacturers as well.
The script outputs the results in an Excel file in the same directory as the project's files.

## Installation
In order to run the script, you need to install Pandas and openpyxl libraries.
Put all the files in the same directory and run the script.


## Usage
When you run the script, there are on-screen instructions on how the model works. In every step, you input the data (#nodes, flow rates, dimensions etc) for the calculations.



## Contributing
You are welcome to give me ideas on how to improve or make my code more versatile without compromising its initial purpose :)
