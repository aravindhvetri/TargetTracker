interface IListNames {
  CRMProjects: string;
  CRMBillings: string;
  TargetConfiguration: string;
}

interface IPeoplePickerDetails {
  id: number;
  name: string;
  email: string;
}

interface ITargetConfigurationData {
  ID?: number | null;
  Technology: string;
  Target: string;
  ProjectManager: IPeoplePickerDetails[];
  DeliveryHead: IPeoplePickerDetails[];
  IsDelete?: boolean;
}
