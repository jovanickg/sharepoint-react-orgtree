import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOrgTreeProps {
  // ... existing props ...
  listTitle: string;
  colTitle: string;
  colJob: string;
  colDept: string;
  colSuperior: string;
  colEmail: string;
  colMobile: string;
  
  // NEW: Job Rank / Sorting
  colJobRank: string; // <--- ADD THIS LINE

  // Contract Logic
  colContractType: string;
  contractTypeFilter: string;

  // ... existing props ...
  webPartWidth: number;
  transparentBackground: boolean;
  context: WebPartContext;
}