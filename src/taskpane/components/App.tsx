import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { IconButton, Button } from "@fluentui/react/lib/Button";
import { TextField } from "@fluentui/react";
// import { Button } from "@fluentui/react-button";
/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  dog: string;
  diffValue: string;
}
const App: React.FC<AppProps> = ({ title, isOfficeInitialized }) => {
  const [state, setState] = React.useState<AppState>({
    dog: "",
    listItems: [],
    diffValue: "",
  });

  React.useEffect(() => {
    setState({
      ...state,
      listItems: [
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
      ],
    });
  }, [setState]);

  const handleChange = (e: any) => {
    setState((prevState) => ({
      ...prevState,
      diffValue: e.target.value,
    }));
    // setState(prevItem => { ...prevItem, surname: e.target.value });
  };

  const clickOn = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";
        range.values = [["hello world"]];

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  const clickOnDifferentFunction = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.worksheets.getActiveWorksheet();
        const ranger = range.getRange("A1:A2:A3:A4");
        const datesz = `${new Date().getUTCDate()}/${new Date().getMonth()}/${new Date().getFullYear()}`;

        // Read the range address
        // range.load("address");
        // const xxx = range.address;
        ranger.values = [[23232], [datesz], ["Heloooooooo"], [234]];
        const ranger2 = range.getRange("F:F").getUsedRange().getLastRow().getOffsetRange(1, 0);
        ranger2.values = [[state.diffValue]];
        await context.sync();
        // console.log(`The range address was ${range}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  const yetAnotherFunction = async () => {
    Excel.run(async (context) => {
      const wss = context.workbook.worksheets.getActiveWorksheet();
      const range = wss.getRange("A1:D5");
      range.load("values");
      range.load("columnCount");
      await context.sync();

      const newR = range.values.map((x) => {
        return x.map((y) => {
          return "Hello==" + y;
        });
      });
      range.values = newR;
      console.log("range.values", range.values);
      console.log("range.columnCount", range.columnCount);
      return context;
    });
  };

  if (!isOfficeInitialized) {
    return (
      <div>
        <Progress
          title={title}
          logo={require("./../../../assets/dogge.png")}
          message="Please sideload your addin to see app body."
        />
      </div>
    );
  }
  return (
    <div>
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/dogge.png")} title={title} message="WelcomeTOExcel" />
        <HeroList message="Discover" items={state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button className="ms-welcome__action" onClick={clickOn}>
            <b>Run</b>
          </Button>
        </HeroList>
        <Button appearance="primary" shape="rounded" onClick={clickOnDifferentFunction}>
          Write to cells
        </Button>
        <Button appearance="primary" shape="rounded" onClick={yetAnotherFunction}>
          Read From cells
        </Button>
        <TextField value={state.diffValue} onChange={handleChange} />
        {/* onChange=
        {(e) => {
          handleChange(e);
        }} */}
        {/* <input placeholder="sdsds" type="text" value={state.diffValue} title="yes">
          yinyang
        </input> */}
      </div>
    </div>
  );
};
export default App;
