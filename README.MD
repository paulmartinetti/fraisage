This repository is a CNC milling aid, likely most useful for home-made automation and robotic machining. The goal is to help the user organize and combine G code blocks of specific cuts for several cuts in the same piece of material, using variables and sub-programs in Fanuc CNC.

This javascript code is to be placed in the script editor in Google sheets. 
From sheet2 onward, a G code is kept for each type of cut. On sheet1, the user specifies the number of desired cuts, cut types, and parameters. The functions assist in the selection and layout of parameters and then provide a single, clean G code text to make all of the cuts.

The G code for the individual cuts is sample only. There are loads of sites where G code for every shape can be found. This program is designed to string several of these together to help reduce repetitive machining.