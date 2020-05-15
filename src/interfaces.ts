import { IFieldInfo } from '@pnp/sp/fields';

export interface ITerm {
  CustomSortOrder: any;
  Id: string;
  IsAvailableForTagging: boolean;
  IsDeprecated: boolean;
  IsRoot: boolean;
  Labels: {
    _ObjectType_: string;
    _Child_Items_: {
      _ObjectType_: string;
      _ObjectIdentity_: string;
      IsDefaultForLanguage: boolean;
      Language: number;
      Value: string;
    }[];
  };
  LocalCustomProperties: any;
  Name: string;
  Parent: {
    _ObjectType_: string;
    _ObjectIdentity_: string;
    CreatedDate: string;
    Id: string;
    LastModifiedDate: string;
    Name: string;
    CustomProperties: any;
    CustomSortOrder: any;
    IsAvailableForTagging: boolean;
    Owner: string;
    Description: string;
    IsDeprecated: boolean;
    IsKeyword: boolean;
    IsPinned: boolean;
    IsPinnedRoot: boolean;
    IsReused: boolean;
    IsRoot: boolean;
    IsSourceTerm: boolean;
    LocalCustomProperties: any;
    MergedTermIds: any[];
    PathOfTerm: string;
    TermsCount: number;
  };
  PathOfTerm: string;
  TermsCount: number;
  _ObjectIdentity_: string;
  _ObjectType_: string;
}

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

export interface Field extends IFieldInfo {
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
