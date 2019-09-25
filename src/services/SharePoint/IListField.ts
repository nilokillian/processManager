export interface IListField {
  Id: string;
  Title: string;
  InternalName: string;
  TypeAsString: string;
  LookupField: string;
  LookupList?: string;
  AllowMultipleValues?: boolean;
  [index: string]: any;
}

export interface IListFieldCollection {
  value: IListField[];
}
