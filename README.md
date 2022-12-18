<a name="readme-top"></a>
![GitHub last commit (branch)](https://img.shields.io/github/last-commit/azmi-maz/inventory-system-for-biochem/main)

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/azmi-maz/inventory-system-for-biochem">
    <img src="https://user-images.githubusercontent.com/87229604/208306950-c85c5315-9ebf-4991-9ff7-fe0b83cad68a.gif" alt="Logo" width="80" height="80">
  </a>

<h3 align="center">Biochemistry Inventory System</h3>

  <p align="center">
    <br />
    <a href="https://github.com/azmi-maz/inventory-system-for-biochem"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://docs.google.com/spreadsheets/d/1DA_8fUuL4t9OM61inJ-E4ElOMapTTt7QdZHvUHDeBkk/edit?usp=share_link">View Demo</a>
  </p>
</div>


<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#installation">Installation</a></li>
        <li><a href="#google-sheets-needed">Google sheets needed/a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#license">License</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

This inventory system was made to handle our stock-in and stock-out processes faster using GS1 data matrix of the Abbott Alinity reagents and consumables.

Main features:
* Dashboard to see in-stock reagents, recent stock transactions, expired reagents, pending purchase requested reagents, and below-par-level reagents.
* Calculates the reorder quantity automatically based on the stock transaction tables to prevent understocking and reduce overstocking.
* Produces statistical reports for monthly and annual reports.
* Facilitates communication by producing reports between laboratory users, the procurement team, and suppliers on active purchase orders.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



### Built With

* ![JavaScript](https://img.shields.io/badge/javascript-%23323330.svg?style=for-the-badge&logo=javascript&logoColor=%23F7DF1E)
* ![Google Drive](https://img.shields.io/badge/Google%20Drive-4285F4?style=for-the-badge&logo=googledrive&logoColor=white)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- GETTING STARTED -->
## Getting Started

This project must be accompanied by another google sheet to allow external users to view/ interact with the inventory system.
<br />

<a href="https://github.com/azmi-maz/supplier-view-inventory-system"><strong>See repo here»</strong></a>

### Installation

1. Create a google sheet
2. Install Google Apps Script GitHub Assistant extension
3. Clone this repo
4. Login to GitHub using the extension with you GitHub token
5. Pull the main branch

### Google sheets needed

*limited - User interacts with protected sheet
<br />
**restricted - hidden sheet

| Sheet Name | User Restriction |
| ----------- | ----------- |
| Dashboard | * |
| INCOMING |  |
| OUTGOING |  |
| MANUAL |  |
| Verification | | 
| IN LIST | * |
| OUT LIST | * |
| QOH PR | * |
| QOH FOC | * |
| tblStockIN | ** |
| tblUniqueINID | ** |
| tblStockOUT | ** |
| Store_Alinity | * |
| EXPIRED | * |
| Cold_Items | * |
| Store-PR | * |
| Store-FOC | * |
| BO-QOH<MTH | * |
| To PR |  |
| Order PR |  |
| Ves-Correct | *
| VExcel | 
| Order FOC | 
| PO Entry | 
| DO Entry | 
| Batch List | 
| tblPR | ** |
| tblPO | ** |
| tblDO | ** |
| MasterL | ** |
| tblBatch | ** |
| tblBestExp | ** |
| Statistics | * |
| TestCountEntry | | 
| Alinity_Count |  |
| FOC-NonAbbott |  |
| CountComp | ** |
| Test_Count Data | ** |
| FOCNonAbbott Data | ** |
| FOCNonAbbottComp | ** |
| Pivots | ** |
| ItemCodeL | ** |
| Correct REQ | * |

Note: to document formulas used in each sheets


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- USAGE EXAMPLES -->
## Usage

INCOMING - scan the item barcode one row at a time and click the "Add" button.
<br />
OUTGOING - Choose the item by checking the checkbox and click "Stock out" button.

Note: to continue documentation as user guide.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## Roadmap

- [ ] Stock-in and stock-out processes can be done with barcding or manually (for items that are not barcoded).
- [ ] Query quantity on hand for purchase request items and FOC.
- [ ] Reagent verification list is auto-populated during stock-in process.
- [ ] Produce department stock requistion forms.
- [ ] Purchase requests and order features:
    - [ ] Query list of received items to ease data entry for user.
    - [ ] Query critical items that are below par level.
    - [ ] Autocalculates reorder stock quantity for need-to-order items.
- [ ] Statistics:
    - [ ] Test count to calculate the test usage - average per month (APM).
    - [ ] Finance data, for monthly and annual report.
- [ ] Dashboard to keep track of the important stuff.

<!-- See the [open issues](https://github.com/azmi-maz/inventory-system-for-biochem/issues) for a full list of proposed features (and known issues). -->

<p align="right">(<a href="#readme-top">back to top</a>)</p>




<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

