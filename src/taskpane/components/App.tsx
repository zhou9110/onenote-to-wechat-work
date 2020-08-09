import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
// import Header from "./Header";
import Progress from "./Progress";
import DateList, { DateListItem } from "./DateList";

/* global Button, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  dates: DateListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      dates: []
    };
  }

  click = async () => {
    /**
     * Insert your OneNote code here
     */
    try {
      await OneNote.run(async context => {
        // Get the current page.
        var page = context.application.getActivePage();

        // Load the page content, get the dates from the table
        page.load("title,contents");

        await context.sync();

        const outline = page.contents.items[0].outline;
        outline.load();
        outline.paragraphs.load();
        await context.sync();

        const table = outline.paragraphs.items[0].table;
        table.load();
        await context.sync();

        const rows = table.rows;
        rows.load();
        await context.sync();

        const rowItems = rows.items;
        const tableRows = rowItems.map(tableRow => tableRow.load());
        tableRows.forEach(tableRow => tableRow.cells.load());
        const firstTwoCols = tableRows.map(tableRow => ({
          date: tableRow.cells.getItemAt(0).load(),
          tasks: tableRow.cells.getItemAt(1).load()
        }));
        const firstTwoColContents = firstTwoCols.map(({ date, tasks }) => {
          return {
            date: date.paragraphs.load(),
            tasks: tasks.paragraphs.load()
          };
        });
        await context.sync();

        const formatted = firstTwoColContents.map(({ date, tasks }) => {
          date.getItemAt(0).load();
          tasks.getItemAt(0).load();
          const date2 = date.getItemAt(0).richText.load();
          const tasks2 = tasks.items.map(item => item.richText.load());
          return { date: date2, tasks: tasks2 };
        });

        await context.sync();

        const dates = formatted.map(({ date, tasks }, index) => {
          return {
            key: `${index}`,
            primaryText: date.text,
            tasks: tasks.map(({ id, text }) => ({
              id,
              text
            }))
          };
        });

        // Queue a command to add an outline to the page.
        this.setState({ dates });

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
      });
    } catch (error) {
      console.trace("Error: " + error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        {/* <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" /> */}
        <DateList message={title} items={this.state.dates}>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            获取日期
          </Button>
        </DateList>
      </div>
    );
  }
}
