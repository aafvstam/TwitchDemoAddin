import * as React from "react";

import { Button, ButtonType } from "office-ui-fabric-react";
import HeroList, { HeroListItem } from "../components/HeroList";

function HomePage() {
  let listItems: HeroListItem[] = [
    {
      icon: "Ribbon",
      primaryText: "Achieve more with Office integration"
    },
    {
      icon: "Unlock",
      primaryText: "Unlock features and functionality"
    },
    {
      icon: "Design",
      primaryText: "Create and visualize like a pro"
    }
  ];

  async function click() {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello About", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      let hostInfoResult = getHostInfo();

      // insert a paragraph at the end of the document.
      const paragraphHostInfo = context.document.body.insertParagraph(hostInfoResult, Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraphHostInfo.font.color = "Red";

      await context.sync();
    });
  }

  function getHostInfo() {
    var _requirements = Office.context.requirements;
    var types = ["Excel", "Word"];
    var minVersions = ["Preview", "1.6", "1.5", "1.4", "1.3", "1.2", "1.1", "1.0"]; // Start with the highest version

    // Loop through types and minVersions
    for (var type in types) {
      for (var minVersion in minVersions) {
        // Append "Api" to the type for set name, i.e. "ExcelApi" or "WordApi"
        if (_requirements.isSetSupported(types[type] + "Api", minVersions[minVersion])) {
          return minVersions[minVersion];
        }
      }
    }

    return "Nothing";
  }

  return (
    <div className="ms-welcome">
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <Button
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: "ChevronRight" }}
          onClick={click}
        >
          Run
        </Button>
      </HeroList>
    </div>
  );
}

export default HomePage;
