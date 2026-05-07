export namespace Config {
  export const ListNames: IListNames = {
    CRMProjects: "CRMProjects",
    CRMBillings: "CRMBillings",
    TargetConfiguration: "TargetConfiguration",
  };

  export const TargetConfigurationData: ITargetConfigurationData = {
    ID: null,
    Technology: "",
    FinancialYear: "",
    Target: "",
    ProjectManager: [],
    DeliveryHead: [],
    IsDelete: false,
  };
}
