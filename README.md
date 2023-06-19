These are Power Shell Scripts that are invoked to Extract, Transform and Load (ETL) data from OpenMUC DLMS data collection process. There are also Excel and Power BI reports. 

At the moment there are the following Scripts: 

1. OpenMUCProcessScript - invoke OpenMUC with different configurations (channels.xml), to avoid concurrent acces to the communications ports
   (TBD - currently Linux scheduler, with fixed scheduling)
2. OpenMUCDataScript - Transform data from openmuc/data/ascii to the requested format
3. OpenMUCProfileProcess - Transform Load Profiles data from OpenMuc to format that can be processed by OpeneMUCData script
4. ExcelScript - load excel workbook with transformed data
5. OpenMUC_DataControl - controls data for time gaps and eliminate duplicates
6. OpenMUCLOADPBI - loads PowerBi report wih data
7. OpemMUCMain - script that runs seesion of the scripts


