export interface IListView {
  ListViewXml: string;
  [index: string]: any;
}

export interface IListItemCollection {
  value: IListView[];
}
