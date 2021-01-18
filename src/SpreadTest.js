/* eslint-disable react-hooks/exhaustive-deps */
import React, { useState, useRef } from "react";
import styled from "styled-components";
import { saveAs } from "file-saver";

import { SpreadSheets } from "@grapecity/spread-sheets-react";
// import "Components/CardView/Spreadsheet/gc.spread.sheets.excel2013white.13.2.1.css";
import "@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.css";
import spreadExcel from "@grapecity/spread-excelio";

//DESIGNER ERROR OCCURS WHEN UNCOMMENTING THE IMPORT BELOW
// import { Designer } from "@grapecity/spread-sheets-designer-react";
//DESIGNER ERROR OCCURS WHEN UNCOMMENTING THE IMPORT ABOVE

// import "@grapecity/spread-sheets-designer-resources-en";
// import "@grapecity/spread-sheets-designer/styles/gc.spread.sheets.designer.min.css";
import GC from "@grapecity/spread-sheets";
if (window.location.origin.includes("test.tradir.io")) {
  GC.Spread.Sheets.LicenseKey =
    "E499651173151541#B0vouHalFTUFx4UK5WQwtCeIN7S5Z5QhVnTQtSbCRXWkhmZURWW9xUMw3SWxQnMrM5N8o6QW5UMTF4d8JmYCZVZUVEeHVWOwdFNKR7RUZXTvoETsV6dIhDWqREdEJXOVRUN5AXUxdDTvREcPZjRvZVMD3icyFVeN5kah3WSIJXdoZlYDV4LzlmQud6LtFkRsNUWzg7SVp7Z0ZWcWFjcwVjYKlHbMJETD96Y7UDe9IWRZRXbhpXNi9mMZNWNmBjdPh6SSJ7RBVHV82ydUplSyEFeCFmSVBlaLlzRzIlVihkMpZkUxF4Lk9WZ8l5ZiojITJCLiU4N9IkNygzNiojIIJCLxQDM5IjMwcTM0IicfJye&Qf35VfiUURJZlI0IyQiwiI4EjL6BCITpEIkFWZyB7UiojIOJyebpjIkJHUiwiI7IDMycDMgMTMxATMyAjMiojI4J7QiwiIzEDNwEjMwIjI0ICc8VkIsISkwqegYyuI0ISYONkIsUWdyRnOiwmdFJCLiEDN5ETNxMzNxETN6kTO4IiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPBZ5SCFndlVWWYB7Z7pHMhJWUENkSld6Ul5URUBnN8E4aUFkRwlDNqBlbWdzS7ZDWKd4YLpXbut4QUlnQahDVIBzYK3UYKRDRxN7VhFWeTFzSjJ4Z5k5ZDJjbvFDWVlDRy34QzE4SBVFvmzl";
  spreadExcel.LicenseKey =
    "E499651173151541#B0vouHalFTUFx4UK5WQwtCeIN7S5Z5QhVnTQtSbCRXWkhmZURWW9xUMw3SWxQnMrM5N8o6QW5UMTF4d8JmYCZVZUVEeHVWOwdFNKR7RUZXTvoETsV6dIhDWqREdEJXOVRUN5AXUxdDTvREcPZjRvZVMD3icyFVeN5kah3WSIJXdoZlYDV4LzlmQud6LtFkRsNUWzg7SVp7Z0ZWcWFjcwVjYKlHbMJETD96Y7UDe9IWRZRXbhpXNi9mMZNWNmBjdPh6SSJ7RBVHV82ydUplSyEFeCFmSVBlaLlzRzIlVihkMpZkUxF4Lk9WZ8l5ZiojITJCLiU4N9IkNygzNiojIIJCLxQDM5IjMwcTM0IicfJye&Qf35VfiUURJZlI0IyQiwiI4EjL6BCITpEIkFWZyB7UiojIOJyebpjIkJHUiwiI7IDMycDMgMTMxATMyAjMiojI4J7QiwiIzEDNwEjMwIjI0ICc8VkIsISkwqegYyuI0ISYONkIsUWdyRnOiwmdFJCLiEDN5ETNxMzNxETN6kTO4IiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPBZ5SCFndlVWWYB7Z7pHMhJWUENkSld6Ul5URUBnN8E4aUFkRwlDNqBlbWdzS7ZDWKd4YLpXbut4QUlnQahDVIBzYK3UYKRDRxN7VhFWeTFzSjJ4Z5k5ZDJjbvFDWVlDRy34QzE4SBVFvmzl";
} else {
  //spread 13 license key
  // GC.Spread.Sheets.LicenseKey =
  //   "tradir.io,397451619435995#B0UaMVlZkFHUwM6bJN6YNF5ZQFVbjdlZyQjUyx6a5VEMM3WcodXWRRTdjRTQQhjVjhEdVdHTxp6LzIjRQhVMPFFWiR4RLhmRxk6U8MGMk94SDljWkJjbJl5ZZVHclFmU6lTV6Z6KWNEbpdnRkNjTqdUSjlldnBnUs9GMBFzcvpENLF7Lx26NjpHMwBHZzBHT5I7TRNjMGBFRRFkWVlUT8Mnark5MlhnZONkN09UVWBXb9sGZRhXaShDNzJ6SLhVM5gnV8YUNtJHUD5mcLdlcxRTc8M5Y65mNEtCUPtkZ8gUTUdzbGhVZiojITJCLiYzQGRkMwUkI0ICSiwyN9MzNxMjM6ATM0IicfJye#4Xfd5nIV34M6IiOiMkIsIyMx8idgMlSgQWYlJHcTJiOi8kI1tlOiQmcQJCLiYjM6AzNwACOxITMwIDMyIiOiQncDJCLi2WauIXakFmc4JiOiMXbEJCLiEJsqHImsLiOiEmTDJCLiUTO9UzM4kTM6ETN4cTOzIiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPn3GUxB7dnB7d0NXSyc4NzIneYlTdUdnMrlGR6MWeGRzU4JXVmFneIJlQPJmeVF5ThBTd6g7N5w4Q6EGZ5gXOyhDcURzVZ9ENRJWSn5UU4IUYLBlQ5FXZrk6RYNDelJjNNxEM626VjVTQzYzU8sYSMo";
  // spreadExcel.LicenseKey =
  //   "tradir.io,397451619435995#B0UaMVlZkFHUwM6bJN6YNF5ZQFVbjdlZyQjUyx6a5VEMM3WcodXWRRTdjRTQQhjVjhEdVdHTxp6LzIjRQhVMPFFWiR4RLhmRxk6U8MGMk94SDljWkJjbJl5ZZVHclFmU6lTV6Z6KWNEbpdnRkNjTqdUSjlldnBnUs9GMBFzcvpENLF7Lx26NjpHMwBHZzBHT5I7TRNjMGBFRRFkWVlUT8Mnark5MlhnZONkN09UVWBXb9sGZRhXaShDNzJ6SLhVM5gnV8YUNtJHUD5mcLdlcxRTc8M5Y65mNEtCUPtkZ8gUTUdzbGhVZiojITJCLiYzQGRkMwUkI0ICSiwyN9MzNxMjM6ATM0IicfJye#4Xfd5nIV34M6IiOiMkIsIyMx8idgMlSgQWYlJHcTJiOi8kI1tlOiQmcQJCLiYjM6AzNwACOxITMwIDMyIiOiQncDJCLi2WauIXakFmc4JiOiMXbEJCLiEJsqHImsLiOiEmTDJCLiUTO9UzM4kTM6ETN4cTOzIiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPn3GUxB7dnB7d0NXSyc4NzIneYlTdUdnMrlGR6MWeGRzU4JXVmFneIJlQPJmeVF5ThBTd6g7N5w4Q6EGZ5gXOyhDcURzVZ9ENRJWSn5UU4IUYLBlQ5FXZrk6RYNDelJjNNxEM626VjVTQzYzU8sYSMo";
  //spread 14 license key
  GC.Spread.Sheets.LicenseKey =
    "tradir.io,261274735935899#B0Emnzo5LkRDNyF6d4F5RJNVcJNnMzhVZ9VFaNdldzBXZyJ6czpFWPdEaTFVRPlkTIlXcVV6bMdGNxBDW4gUcKJFTUBHUrUVO42GS9gWNvBjQStidsJFZrh4NZ3CSJp5Rap7b5dTNxU4c52WNkhVSQtmYLBVbYNXeaRHSXZXNZN4Q7MGRrNXNPFWdjpXbk5EMX3Ca5Z6Q7YEMndXbNRGU8U6VUpWSnhDS0NDRTJDRzIVSQF6bxU6Y9RTNzMjdBVTYQh6LZRUVaR7c4kmMzBlNxJVW8NGTOJ4SplXQuhmUDpWMaZkTPpkI0IyUiwiIwIDNGVkRDRjI0ICSiwiM7ETO9UDN5gTM0IicfJye&Qf35VfiUURJZlI0IyQiwiI4EjL6BCITpEIkFWZyB7UiojIOJyebpjIkJHUiwiIwUDOxQDMggTMxATMyAjMiojI4J7QiwiIvlmLylGZhJHdiojIz5GRiwiIRCr0BiJ1iojIh94QiwiI9kDO5MTO5MzN4cjMxYjMiojIklkIs4XZzxWYmpjIyNHZisnOiwmbBJye0ICRiwiI34TUD5GUqNDe4FVeLh6c45URPVWVvcXdxVVdvhTMxtidtdUUX56KZNnYzV4aNF5aoNnR85WOpVnayEkU7YjRUFGdyo6KOh6TFpnVlhGe7oWUBN7ZpdkSydUQXd7Nq3EOJNmb0p7M6gFNK3SWroVMQDE20";
  spreadExcel.LicenseKey =
    "tradir.io,261274735935899#B0Emnzo5LkRDNyF6d4F5RJNVcJNnMzhVZ9VFaNdldzBXZyJ6czpFWPdEaTFVRPlkTIlXcVV6bMdGNxBDW4gUcKJFTUBHUrUVO42GS9gWNvBjQStidsJFZrh4NZ3CSJp5Rap7b5dTNxU4c52WNkhVSQtmYLBVbYNXeaRHSXZXNZN4Q7MGRrNXNPFWdjpXbk5EMX3Ca5Z6Q7YEMndXbNRGU8U6VUpWSnhDS0NDRTJDRzIVSQF6bxU6Y9RTNzMjdBVTYQh6LZRUVaR7c4kmMzBlNxJVW8NGTOJ4SplXQuhmUDpWMaZkTPpkI0IyUiwiIwIDNGVkRDRjI0ICSiwiM7ETO9UDN5gTM0IicfJye&Qf35VfiUURJZlI0IyQiwiI4EjL6BCITpEIkFWZyB7UiojIOJyebpjIkJHUiwiIwUDOxQDMggTMxATMyAjMiojI4J7QiwiIvlmLylGZhJHdiojIz5GRiwiIRCr0BiJ1iojIh94QiwiI9kDO5MTO5MzN4cjMxYjMiojIklkIs4XZzxWYmpjIyNHZisnOiwmbBJye0ICRiwiI34TUD5GUqNDe4FVeLh6c45URPVWVvcXdxVVdvhTMxtidtdUUX56KZNnYzV4aNF5aoNnR85WOpVnayEkU7YjRUFGdyo6KOh6TFpnVlhGe7oWUBN7ZpdkSydUQXd7Nq3EOJNmb0p7M6gFNK3SWroVMQDE20";
  GC.Spread.Sheets.Designer.LicenseKey =
    "tradir.io,423825181336266#B0rzsXOz4keP9mYkhlVW3COPVXQYVXeHRWevIGVp36b9UTMWJzQhlWbWFnYZNHVDVjNC9GTsp7anpXSC3yKLx4dIJUMrUXNpJXVI9WZkZDNroHOKpUTxg4YFBDMiVEeoJUWlZEVKFXNFZmQygVOux4R7cjRyUTc9R6Lu5mekhnRwkWbUdTOHFmb5VleyhXOkhmYNJFc9Z4LnhlVrRXVuh7SOZFTNJzUzUGZvAVNGFFdDF5QZ3CR5IGbWFFZxBncCl6dZlGcvJUN5xmQ4R6YORDUNhVTLpVUpB7Lr3iZNBDW4YEWBJiOiMlIsISREJERwIEN7IiOigkIscTO7ETOyAzM0IicfJye&Qf35VfiklNFdjI0IyQiwiI4EjL6BibvRGZB5icl96ZpNXZE5yUKRWYlJHcTJiOi8kI1tlOiQmcQJCLiMTN8EDNwACOxEDMxIDMyIiOiQncDJCLi2WauIXakFmc4JiOiMXbEJCLiEJsqHImsLiOiEmTDJCLiYjNyYzMzEDOxUjM8MjM4IiOiQWSisnOiQkIsISP3EUSWZEcuhXUNdWbLp4YmhmbOBjQChUWzk5V52iVVNHVBJ6ctFmW4EGbadXNoZkWKtUTrc5ZLRER75mRlZWdwtyYGlTTwM7Krw6QvE7cMpHeIRnMhZFe8QDdyskcU3GWoBzY7REe7Q6boB5TFhkbUFDLZuj";
}

// GC.Spread.Sheets.LicenseKey =
//   "tradir.io,397451619435995#B0UaMVlZkFHUwM6bJN6YNF5ZQFVbjdlZyQjUyx6a5VEMM3WcodXWRRTdjRTQQhjVjhEdVdHTxp6LzIjRQhVMPFFWiR4RLhmRxk6U8MGMk94SDljWkJjbJl5ZZVHclFmU6lTV6Z6KWNEbpdnRkNjTqdUSjlldnBnUs9GMBFzcvpENLF7Lx26NjpHMwBHZzBHT5I7TRNjMGBFRRFkWVlUT8Mnark5MlhnZONkN09UVWBXb9sGZRhXaShDNzJ6SLhVM5gnV8YUNtJHUD5mcLdlcxRTc8M5Y65mNEtCUPtkZ8gUTUdzbGhVZiojITJCLiYzQGRkMwUkI0ICSiwyN9MzNxMjM6ATM0IicfJye#4Xfd5nIV34M6IiOiMkIsIyMx8idgMlSgQWYlJHcTJiOi8kI1tlOiQmcQJCLiYjM6AzNwACOxITMwIDMyIiOiQncDJCLi2WauIXakFmc4JiOiMXbEJCLiEJsqHImsLiOiEmTDJCLiUTO9UzM4kTM6ETN4cTOzIiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPn3GUxB7dnB7d0NXSyc4NzIneYlTdUdnMrlGR6MWeGRzU4JXVmFneIJlQPJmeVF5ThBTd6g7N5w4Q6EGZ5gXOyhDcURzVZ9ENRJWSn5UU4IUYLBlQ5FXZrk6RYNDelJjNNxEM626VjVTQzYzU8sYSMo";
// GC.Spread.Sheets.LicenseKey =
//   "E395966134557535#B04ZWROZ5dUFDRZp6MnZVT4wmdrYjT8lVZjxEU844aJlDWU3ia7lmaS5UdaR5Zjxka8ZzcJhzdwMUYtdHaoB7RaB7UTNUdTZlN5kEb5IER9YHMJd4dIZWcFpXcBdTQFBnbO34UyBzaIJTaqpXT94UeB9kR42GeKx6LshUOX54a6FmZ0Z5L92EW6RkYwgFNsZVYtxGNpl4V4sUTk3WN82kQMdUNshjdr94K0ZkV4lENKJjcGVXZBpkRBplSUJWNzYkW7omTTNTWsV7YUN4SGJ4SwgVcRhzdrQDcjhjND5WTTZ4T84WbFRGZWV7STJiOiMlIsIiRGZDOFVkQiojIIJCL4gDOwkzM5UjM0IicfJye#4Xfd5nIV34M6IiOiMkIsIyMx8idgMlSgQWYlJHcTJiOi8kI1tlOiQmcQJCLigTNwMjNwASMyITMwIDMyIiOiQncDJCLiEjMzATMyAjMiojIwhXRiwiIRCr0BiJ1iojIh94QiwSZ5JHd0ICb6VkIsISNzUzN5UDNzEjN6kTN9MjI0ICZJJCL3V6csFmZ0IiczRmI1pjIs9WQisnOiQkIsISP3E5b7ZVYvsGOZlWNZhUQ0ZVM93GUwsyLK96K7hzU4l6NZVlcUBTV734bilVbyEVUGBDR4ske8k5ZnRXNq36R7plbrE5KZBlNVZHbTZFMsNGepdkZ5FTT8JXaIdkNCVERuBTSCRDTTdkRKY8c";
// GC.Spread.Sheets.Designer.LicenseKey =
//   "tradir.io,397451619435995#B0UaMVlZkFHUwM6bJN6YNF5ZQFVbjdlZyQjUyx6a5VEMM3WcodXWRRTdjRTQQhjVjhEdVdHTxp6LzIjRQhVMPFFWiR4RLhmRxk6U8MGMk94SDljWkJjbJl5ZZVHclFmU6lTV6Z6KWNEbpdnRkNjTqdUSjlldnBnUs9GMBFzcvpENLF7Lx26NjpHMwBHZzBHT5I7TRNjMGBFRRFkWVlUT8Mnark5MlhnZONkN09UVWBXb9sGZRhXaShDNzJ6SLhVM5gnV8YUNtJHUD5mcLdlcxRTc8M5Y65mNEtCUPtkZ8gUTUdzbGhVZiojITJCLiYzQGRkMwUkI0ICSiwyN9MzNxMjM6ATM0IicfJye#4Xfd5nIV34M6IiOiMkIsIyMx8idgMlSgQWYlJHcTJiOi8kI1tlOiQmcQJCLiYjM6AzNwACOxITMwIDMyIiOiQncDJCLi2WauIXakFmc4JiOiMXbEJCLiEJsqHImsLiOiEmTDJCLiUTO9UzM4kTM6ETN4cTOzIiOiQWSiwSflNHbhZmOiI7ckJye0ICbuFkI1pjIEJCLi4TPn3GUxB7dnB7d0NXSyc4NzIneYlTdUdnMrlGR6MWeGRzU4JXVmFneIJlQPJmeVF5ThBTd6g7N5w4Q6EGZ5gXOyhDcURzVZ9ENRJWSn5UU4IUYLBlQ5FXZrk6RYNDelJjNNxEM626VjVTQzYzU8sYSMo";
const CardViewSpreadsheet = () => {
  const attachRef = useRef();
  const [importExcelFile, setImportExcelFile] = useState(null);
  const [exportFileName, setExportFileName] = useState("newAttachment.xlsx");
  const [password, setPassword] = useState("");

  const changeFileDemo = (e) => {
    setImportExcelFile(e.target.files[0]);
  };
  const changePassword = (e) => {
    setPassword(e.target.value);
  };
  const changeExportFileName = (e) => {
    setExportFileName(e.target.value);
  };
  // const changeIncremental = (e) => {
  //   document.getElementById("loading-container").style.display = e.target
  //     .checked
  //     ? "block"
  //     : "none";
  // };
  const loadExcel = (e) => {
    let excelIo = new spreadExcel.IO();
    let excelFile = importExcelFile;

    // let incrementalEle = document.getElementById("incremental");
    // let loadingStatus = document.getElementById("loadingStatus");
    // here is excel IO API
    excelIo.open(
      excelFile,
      function (json) {
        // let workbookObj = json;
        // if (incrementalEle.checked) {
        //   attachRef.current.spread.fromJSON(workbookObj, {
        //     incrementalLoading: {
        //       loading: function (progress) {
        //         progress = progress * 100;
        //         loadingStatus.value = progress;
        //       },
        //       loaded: function () {},
        //     },
        //   });
        // } else {
        //   attachRef.current.spread.fromJSON(workbookObj);
        // }
        attachRef.current.spread.fromJSON(json);
      },
      function (e) {
        // process error
        alert(e.errorMessage);
      },
      { password: password }
    );
  };
  const saveExcel = (e) => {
    let excelIo = new spreadExcel.IO();

    let fileName = exportFileName;
    if (fileName.substr(-5, 5) !== ".xlsx") {
      fileName += ".xlsx";
    }

    let json = attachRef.current.spread.toJSON();

    // here is excel IO API
    excelIo.save(
      json,
      function (blob) {
        saveAs(blob, fileName);
      },
      function (e) {},
      { password: password }
    );
  };
  const saveToAttachment = (e) => {
    let excelIo = new spreadExcel.IO();

    let fileName = exportFileName;
    if (fileName.substr(-5, 5) !== ".xlsx") {
      fileName += ".xlsx";
    }

    let json = attachRef.current.spread.toJSON();

    // here is excel IO API
    excelIo.save(
      json,
      function (blob) {
        const newFile = new FormData();
        newFile.append("file", blob);
        newFile.append("name", fileName);
        newFile.append("content_type", "xlsx");
      },
      function (e) {},
      { password: password }
    );
  };

  return (
    <CardNoteContainer>
          <FileContainer>
            <FileBox>
                <FileButton onClick={(e) => loadExcel(e)}>import from</FileButton>
              <ImportInput type="file" onChange={(e) => changeFileDemo(e)} />
            </FileBox>
            <FileBox>
              <FileButton onClick={(e) => saveExcel(e)}>download as</FileButton>
              <FileButton onClick={(e) => saveToAttachment(e)}>
                send to shared attachments as
              </FileButton>
              <ExportInput
                defaultValue="newAttachment.xlsx"
                onChange={(e) => changeExportFileName(e)}
              />
            </FileBox>
            <FileBox>
              Set export file pw:&nbsp;
              <ExportInput
                type="password"
                onChange={(e) => changePassword(e)}
              />
            </FileBox>
          </FileContainer>
          <BodyContainer>
              <SpreadContainer>
                <SpreadSheets
                  ref={attachRef}
                  allowUserDragMerge
                  backColor="white"
                  hostStyle={{
                    width: "100%",
                    height: "400px",
                    // marginTop: "20px",
                  }}
                ></SpreadSheets>
              </SpreadContainer>
          </BodyContainer>
    
    </CardNoteContainer>
  );
};

export default CardViewSpreadsheet;

const FileContainer = styled.div`
  width: 100%;
`;

const FileBox = styled.div`
  display: flex;
  width: 100%;
  line-height: 30px;
  padding-bottom: 4px;
`;

const FileButton = styled.button`
  height: 30px;
  margin: 0 5px;
`;

const ExportInput = styled.input`
  width: 300px;
  margin-left: 10px;
`;

const ImportInput = styled.input`
  border: none;
`;

const SpreadContainer = styled.div`
  width: 100%;
`;

const CardNoteContainer = styled.div`
`;

const BodyContainer = styled.div`
  border: 1px solid ${({ theme }) => theme.colors.color_base_dark};
  transition: all 1s ease-in-out;
  overflow: hidden;
  display: flex;
  justify-content: center;
  align-items: center;
  position: relative;
  margin-bottom: 32px;
`;
