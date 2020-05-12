export interface IListFields {
  FieldChoice: FieldChoice;
  FieldComputed: FieldComputed;
  FieldDateTime: FieldDateTime;
  FieldLookup: FieldLookup;
  FieldMultiLineText: FieldMultiLineText;
  FieldNumber: FieldNumber;
  FieldText: FieldText;
  FieldUser: FieldUser;
  TaxonomyField: TaxonomyField;
}

export type IListField = IListFields[keyof IListFields];

export interface Field {
  AutoIndexed: boolean;
  CanBeDeleted: boolean;
  ClientSideComponentId: string;
  ClientSideComponentProperties: any;
  ClientValidationFormula: any;
  ClientValidationMessage: any;
  CustomFormatter: any;
  DefaultFormula: any;
  DefaultValue: string;
  Description: string;
  Direction: string;
  EnforceUniqueValues: boolean;
  EntityPropertyName: string;
  FieldTypeKind: number;
  Filterable: boolean;
  FromBaseType: boolean;
  Group: string;
  Hidden: boolean;
  Id: string;
  IndexStatus: number;
  Indexed: boolean;
  InternalName: string;
  JSLink: string;
  PinnedToFiltersPane: boolean;
  ReadOnlyField: boolean;
  Required: boolean;
  SchemaXml: string;
  Scope: string;
  Sealed: boolean;
  ShowInFiltersPane: number;
  Sortable: boolean;
  StaticName: string;
  Title: string;
  TypeAsString: string;
  TypeDisplayName: string;
  TypeShortDescription: string;
  ValidationFormula: any;
  ValidationMessage: any;
  'odata.editLink': string;
  'odata.id': string;
}

export interface FieldChoice extends Field {
  Choices: string[];
  EditFormat: number;
  FillInChoice: boolean;
  Mappings: any;
  'odata.type': 'SP.FieldChoice';
}

export interface FieldComputed extends Field {
  EnableLookup: boolean;
  'odata.type': 'SP.FieldComputed';
}

export interface FieldDateTime extends Field {
  DateTimeCalendarType: number;
  DisplayFormat: number;
  FriendlyDisplayFormat: number;
  'odata.type': 'SP.FieldDateTime';
}

export interface FieldLookup extends Field {
  AllowMultipleValues: boolean;
  DependentLookupInternalNames: any[];
  IsDependentLookup: boolean;
  IsRelationship: boolean;
  LookupField: string;
  LookupList: string;
  LookupWebId: string;
  PrimaryFieldId: string;
  RelationshipDeleteBehavior: number;
  UnlimitedLengthInDocumentLibrary: boolean;
  'odata.type': 'SP.FieldLookup';
}

export interface FieldMultiLineText extends Field {
  AllowHyperlink: boolean;
  AppendOnly: boolean;
  NumberOfLines: number;
  RestrictedMode: boolean;
  RichText: boolean;
  UnlimitedLengthInDocumentLibrary: boolean;
  WikiLinking: boolean;
  'odata.type': 'SP.FieldMultiLineText';
}

export interface FieldNumber extends Field {
  DisplayFormat: number;
  MaximumValue: number;
  MinimumValue: number;
  ShowAsPercentage: boolean;
  'odata.type': 'SP.FieldNumber';
}

export interface FieldText extends Field {
  MaxLength: number;
  'odata.type': 'SP.FieldText';
}

export interface FieldUser extends Field {
  AllowDisplay: boolean;
  AllowMultipleValues: boolean;
  DependentLookupInternalNames: any[];
  IsDependentLookup: boolean;
  IsRelationship: boolean;
  LookupField: string;
  LookupList: string;
  LookupWebId: string;
  Presence: boolean;
  PrimaryFieldId: any;
  RelationshipDeleteBehavior: number;
  SelectionGroup: number;
  SelectionMode: number;
  UnlimitedLengthInDocumentLibrary: boolean;
  'odata.type': 'SP.FieldUser';
}

export interface TaxonomyField extends Field {
  AllowMultipleValues: boolean;
  AnchorId: string;
  CreateValuesInEditForm: boolean;
  DependentLookupInternalNames: any[];
  IsAnchorValid: boolean;
  IsDependentLookup: boolean;
  IsKeyword: boolean;
  IsPathRendered: boolean;
  IsRelationship: boolean;
  IsTermSetValid: boolean;
  LookupField: string;
  LookupList: string;
  LookupWebId: string;
  Open: boolean;
  PrimaryFieldId: any;
  RelationshipDeleteBehavior: number;
  SspId: string;
  TargetTemplate: any;
  TermSetId: string;
  TextField: string;
  UnlimitedLengthInDocumentLibrary: boolean;
  UserCreated: boolean;
  'odata.type': 'SP.Taxonomy.TaxonomyField';
}
