export interface IChoiceField {
  Choices: string[];
  DefaultValue: string;
  EntityPropertyName: string;
  Id: string;
  InternalName: string;
  TypeAsString: string;
  Title: string;
  [index: string]: any;
}

export interface IChoiceFieldCollection {
  value: IChoiceField[];
}
