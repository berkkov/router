# src
Contributors: Ipek Erdil, Busra Karkili, Emre Sarigedik, Goksu Ureten, Eti Food Co.
Partly powered by Trafi

A practical tool for:
-Clustering customers by their locations and additional retail related parametrers
-Scheduling and routing the workforce via public transportation means of a city between the customers

cluster.py
Clusters customers by their locations and either their Monthly Average Revenue Balance or pre-defined service times.

MapClusters.py
Creates an interactive map of customer groups generated after cluster.py

ModelStrict.py and model2.py
MIP models for routing and scheduling.

end_module.py
Repetetively solves small MIP models created above for every cluster.

ScheduleWriter.py
Results of the above models are not readable. This script creates schedule and writes them on an Excel file.


Requirements:
pandas
pulp
openpyxl
webbrowser
xlsxwriter
xlrd
pathlib
asyncio
idna
numpy
sklearn - not a real requirement but recommended for service time and visit frequency calculations if you need to

In case I missed:
Ciro = Revenue
Musteri = Customer