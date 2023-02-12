
# Price creator

To run the script, you need to install a [NodeJs v.19 or higer](https://nodejs.org/en/download/current/).\
Then execute the command in the script folder:


## Run Locally

Clone the project

```bash
  git clone https://github.com/dimamarjan/new-arrivals-price-builder
```

Go to the project directory

```bash
  cd my-project
```

Install dependencies

```bash
  npm install
```

Start the script

(this command will create an XLSX file in the folder "~/build" with the file name of the current date and the specified prefix in the file "/common/dataVars.js")

```bash
  npm run start 
```


## Features

- Put the download file from TradeCRM in "~/source" folder (named as All.xlsx file from TradeCRM)
    * That file must include columns: "Код товара", "Название","Цена Розничная", "URL Изображения"
- Put the "txt" file into "~/source" with the specified item names separated by "Enter" (each position on a new line).
    * Еhe names of the products in the TXT file must be identical to the names of the products in TradeCRM.


