# Mapped_Chart_Quadrats_VBA_Code
This repository contains VBA and ArcObjects code used to analyze plant distributions in digitized quadrats near Flagstaff Arizona, over the years 2002 - 2020.
This code was used to produce the data presented in the Data Paper "Cover and density of southwestern ponderosa pine understory plants in permanent chart quadrats (2002-2020)" (Moore et al. In Review).

The relevant functions are embedded in larger modules containing other unused functions (17 VBA modules containing 791 functions and 78,912 lines of code).  The primary analytical master function is "RunAsBatch" in the module "ThisDocument_for_VM_2".  This function runs several other functions that do the various steps of the analysis.  The primary map export function is "ExportImages" in the module "Quadrat_Map_Module", and this function creates common plant species symbology that can be applied to all 1,500+ maps, and exports individual maps for each quadrat and for each year.

Moore, M. M., J. S. Jenness, D. C. Laughlin, R. T. Strahan, J. D. Bakker, H. E. Dowling, and J. D. Springer. In Review. Cover and density of southwestern ponderosa pine understory plants in permanent chart quadrats (2002-2020). Ecology. Data Paper.
