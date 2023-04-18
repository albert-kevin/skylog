### generate data

use ChatGPT-4 with the instruction set **agent.txt**.  
use swiftkey with Google voice-to-text to add data to chatbot on mobile phone.  
Ask the bot to print the CSV dataframe text.  

### data input

copy the generated **CSV dataframe** text from ChatGPT-4 and paste it into **input.csv**.  

### Modules

Make sure you installed these modules on at least **Python 3.8**  
**pip install pandas==2.0.0 xlsxwriter==3.1.0**  

### execute 

run the script either using the notebook **report.ipynb** or the script **report.py**.  

output:  
**backup of the input file with the date it was generated**  
/data/bronze/input_backup_18Apr2023  

**for each pilot one file, for example:**  
/data/silver/18Apr2023__Chris__Vluchtlogboek.xlsx  
/data/silver/18Apr2023__Tom__Vluchtlogboek.xlsx  
/data/silver/18Apr2023__Sylvia__Vluchtlogboek.xlsx  

**zip file:**  
/data/silver/18Apr2023_Vluchtlogboeken.zip  

### folder structure

/readme.md  
/LICENSE  
/agent.txt  
/data/bronze/input_backup_18Apr2023  
/data/bronze/input.csv  
/data/silver/18Apr2023__Chris__Vluchtlogboek.xlsx  
/data/silver/18Apr2023__Tom__Vluchtlogboek.xlsx  
/data/silver/18Apr2023__Sylvia__Vluchtlogboek.xlsx  
/data/silver/18Apr2023_Vluchtlogboeken.zip  
/code/scripts/report.py  
/code/notebooks/report.ipynb  