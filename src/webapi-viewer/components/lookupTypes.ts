export interface LookupEntityMetadata {
  logicalName: string;
  entitySetName: string;
  primaryIdAttribute: string;
  primaryNameAttribute?: string | null;
}

export interface LookupSelection {
  logicalName: string;
  entitySetName: string;
  recordId: string;
  displayName: string;
}
