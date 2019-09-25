export interface IList {
  Id: string;
  Title: string;
  ItemCount: number;
  [index: string]: any;
}

export interface IListCollection {
  value: IList[];
}
