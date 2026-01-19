export interface EntityMetadata {
  LogicalName: string;
  DisplayName: string;
  EntitySetName: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  Description?: string;
}

export interface AttributeMetadata {
  LogicalName: string;
  DisplayName: string;
  AttributeType: string;
  IsPrimaryId?: boolean;
  IsPrimaryName?: boolean;
}

export interface AttributeMetadataComplete extends AttributeMetadata {
  RequiredLevel?: string;
  MaxLength?: number;
  Precision?: number;
  Scale?: number;
  Description?: string;
  OptionSetValues?: Array<{ Value: number; Label: string }>;
  LookupTargets?: string[];
  IsCalculated?: boolean;
  IsRollup?: boolean;
  IsPolymorphic?: boolean;
  IsActivityParty?: boolean;
}

export interface RelationshipMetadata {
  SchemaName: string;
  ReferencingEntity: string;
  ReferencedEntity: string;
  ReferencingAttribute: string;
  ReferencingEntityNavigationPropertyName: string;
  ReferencedEntityNavigationPropertyName?: string;
  RelationshipType: 'OneToMany' | 'ManyToOne' | 'ManyToMany';
}

export interface EntityMetadataComplete extends EntityMetadata {
  Attributes: AttributeMetadataComplete[];
  OneToManyRelationships: RelationshipMetadata[];
  ManyToOneRelationships: RelationshipMetadata[];
  ManyToManyRelationships: RelationshipMetadata[];
}

export interface QueryFilter {
  id: string;
  type: 'condition' | 'group';
  logicalOperator?: 'and' | 'or'; // for group
  
  // for condition
  entityAlias?: string; // 'main' or alias of related entity
  attribute?: string;
  operator?: string;
  value?: any;
  value2?: any; // For "Between" or relative date params
  
  // for group
  children?: QueryFilter[];
}

export interface QueryColumn {
  entityAlias: string; // 'main' or alias
  attribute: string;
  displayName: string;
  logicalName?: string;
  attributeType?: string;
}

export interface SelectedEntity {
  logicalName: string;
  entitySetName: string;
  displayName: string;
  primaryIdAttribute: string;
  primaryNameAttribute: string;
  alias: string;
}

export interface JoinedEntity extends SelectedEntity {
  relationshipName: string;
  parentAlias: string;
  relationshipType: 'OneToMany' | 'ManyToOne' | 'ManyToMany';
  navigationPropertyName: string;
}

export interface ViewMetadata {
  id: string;
  name: string;
  fetchXml: string;
  queryType: number;
  isUserQuery: boolean;
}
