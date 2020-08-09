import * as React from "react";
import { Empty, Table } from "antd";

const { Column } = Table;

export interface DateListItem {
  key: string;
  primaryText: string;
  tasks: { id: string; text: string }[];
}

export interface DateListProps {
  message: string;
  items: DateListItem[];
}

export default class DateList extends React.Component<DateListProps> {
  render() {
    const { children, items, message } = this.props;

    return (
      <main className="ms-welcome__main">
        <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2>
        {!items || items.length === 0 ? (
          <Empty />
        ) : (
          <Table dataSource={items.slice(1)} style={{ width: "100%" }}>
            <Column title={items[0].primaryText} dataIndex="primaryText" key="dates"></Column>
            <Column
              title={items[0].tasks[0].text}
              dataIndex="tasks"
              key="tasks"
              render={value => value.map(({ id, text }) => <p key={id}>{text}</p>)}
            />
          </Table>
        )}
        {children}
      </main>
    );
  }
}
