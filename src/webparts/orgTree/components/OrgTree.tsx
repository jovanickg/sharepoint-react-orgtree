import * as React from 'react';
import { IOrgTreeProps } from './IOrgTreeProps';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Tree, TreeNode } from 'react-organizational-chart';
import styles from './OrgTree.module.scss';
import { Dialog, DialogType, DialogFooter, PrimaryButton, Persona, PersonaSize, Icon } from '@fluentui/react';

// INTERFACES
interface IEmployee {
  Id?: number;
  Title: string;
  Job: string;
  Dept: string;
  Superior: string;
  ContractType: string;
  JobRank: string; // NEW: Used for sorting
  Email: { EMail: string };
  Mobile: string; 
}

interface IDepartmentNode {
  id: string;
  name: string;
  employees: IEmployee[];
  contractors: IEmployee[];
  children: IDepartmentNode[];
  isContractorNode?: boolean;
}

interface IOrgTreeState {
  rootNode: IDepartmentNode | undefined;
  isLoading: boolean;
  error: string | undefined;
  zoom: number;
  selectedEmployee: IEmployee | undefined; // For Popup
}

export default class OrgTree extends React.Component<IOrgTreeProps, IOrgTreeState> {

  private containerRef = React.createRef<HTMLDivElement>();
  private scrollAreaRef = React.createRef<HTMLDivElement>();
  private treeWrapperRef = React.createRef<HTMLDivElement>();
  private rootCardRef = React.createRef<HTMLDivElement>();

  // DRAG VARIABLES
  private pos = { top: 0, left: 0, x: 0, y: 0 };

  constructor(props: IOrgTreeProps) {
    super(props);
    this.state = {
      rootNode: undefined,
      isLoading: true,
      error: undefined,
      zoom: 1,
      selectedEmployee: undefined
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  public async componentDidUpdate(prevProps: IOrgTreeProps): Promise<void> {
    const p = this.props;
    const pp = prevProps;
    // Reload if any key mapping or list title changes
    if (pp.listTitle !== p.listTitle || 
        pp.contractTypeFilter !== p.contractTypeFilter ||
        pp.colDept !== p.colDept || 
        pp.colSuperior !== p.colSuperior || 
        pp.colMobile !== p.colMobile ||
        pp.colJobRank !== p.colJobRank) { // Added check for Rank column
      await this.loadData();
    }
  }

  private async loadData(): Promise<void> {
    this.setState({ isLoading: true, error: undefined });

    try {
      if (!this.props.listTitle) {
        this.setState({ isLoading: false });
        return;
      }

      const sp = spfi().using(SPFx(this.props.context));

      const fTitle = this.props.colTitle || "Title";
      const fJob = this.props.colJob || "Job_x0020_Title";
      const fDept = this.props.colDept || "Department";
      const fSup = this.props.colSuperior || "Superior_Department";
      const fContract = this.props.colContractType || "Contract_Type";
      const fEmail = this.props.colEmail || "Email";
      const fMobile = this.props.colMobile || "MobilePhone";
      // NEW: Fetch the rank column (defaults to Job_Position_Code)
      const fRank = this.props.colJobRank || "Job_Position_Code";

      const rawItems = await sp.web.lists
        .getByTitle(this.props.listTitle)
        .items
        .select("Id", fTitle, fJob, fDept, fSup, fContract, fMobile, fRank, `${fEmail}/EMail`)
        .expand(fEmail)
        .top(5000)(); 

      const mappedEmployees: IEmployee[] = rawItems.map((item: any) => ({
        Id: item.Id,
        Title: item[fTitle],
        Job: item[fJob],
        Dept: item[fDept],
        Superior: item[fSup],
        ContractType: item[fContract],
        // Default to "999999" so employees with no rank appear at the bottom
        JobRank: item[fRank] || "999999", 
        Email: item[fEmail],
        Mobile: item[fMobile]
      }));

      const tree = this.buildDepartmentTree(mappedEmployees);

      this.setState({ rootNode: tree || undefined, isLoading: false }, () => {
        setTimeout(this.handleCenterRoot, 300);
      });

    } catch (error) {
      console.error("Error loading Org Tree", error);
      this.setState({ isLoading: false, error: "Failed to load list. Check mappings." });
    }
  }

  private buildDepartmentTree(employees: IEmployee[]): IDepartmentNode | null {
    const deptMap = new Map<string, IDepartmentNode>();
    const validContracts = (this.props.contractTypeFilter || "").split(',').map(s => s.trim().toUpperCase());

    // 1. Group Employees into Departments
    employees.forEach(emp => {
      const deptName = emp.Dept || "Unassigned";
      if (!deptMap.has(deptName)) {
        deptMap.set(deptName, { id: deptName, name: deptName, employees: [], contractors: [], children: [] });
      }
      const node = deptMap.get(deptName)!;
      const empContract = (emp.ContractType || "").toUpperCase();
      if (validContracts.length === 0 || validContracts.indexOf(empContract) !== -1) {
        node.employees.push(emp);
      } else {
        node.contractors.push(emp);
      }
    });

    // 2. Sorting Helper (Low number = Higher Rank)
    const sortEmps = (a: IEmployee, b: IEmployee) => {
        const rA = a.JobRank || "999999";
        const rB = b.JobRank || "999999";
        return rA.localeCompare(rB); // String compare works for fixed width patterns like "0100.01"
    };

    let root: IDepartmentNode | null = null;

    // 3. Build Tree Structure
    deptMap.forEach((node, deptName) => {
      // A. Sort Employees INSIDE the card immediately
      node.employees.sort(sortEmps);
      node.contractors.sort(sortEmps);

      const sampleEmp = node.employees[0] || node.contractors[0];
      if (!sampleEmp) return;
      
      const parentName = sampleEmp.Superior;
      const isRootCandidate = !parentName || parentName === deptName || parentName.trim() === '';
      
      if (isRootCandidate) {
        if (!root || (node.employees.some(e => e.Job && /direktor|ceo|president/i.test(e.Job)))) {
          root = node;
        }
      } else {
        if (deptMap.has(parentName)) {
            const parent = deptMap.get(parentName)!;
            if(parent.children.indexOf(node) === -1) parent.children.push(node);
        }
      }
    });

    // 4. Recursive Function to Sort BRANCHES (Children)
    // We sort sub-departments based on the Rank of their "Boss" (the first employee)
    const sortChildren = (node: IDepartmentNode) => {
        if (!node.children || node.children.length === 0) return;

        node.children.sort((a, b) => {
            // Get the boss of department A
            const bossA = a.employees.length > 0 ? a.employees[0] : (a.contractors[0] || null);
            // Get the boss of department B
            const bossB = b.employees.length > 0 ? b.employees[0] : (b.contractors[0] || null);

            const rankA = bossA ? (bossA.JobRank || "999999") : "999999";
            const rankB = bossB ? (bossB.JobRank || "999999") : "999999";

            return rankA.localeCompare(rankB);
        });

        // Recursively sort grandchildren
        node.children.forEach(child => sortChildren(child));
    };

    if (root) {
        sortChildren(root);
    }

    return root;
  }

  // --- INTERACTION ---
  private handleEmployeeClick = (emp: IEmployee, e: React.MouseEvent): void => {
    e.stopPropagation(); // Prevent drag start
    this.setState({ selectedEmployee: emp });
  }

  private closeDialog = (): void => {
      this.setState({ selectedEmployee: undefined });
  }

  // --- DRAG HANDLERS ---

  private onMouseDown = (e: React.MouseEvent<HTMLDivElement>): void => {
    const ele = this.scrollAreaRef.current;
    if (!ele) return;

    // Only drag if not clicking a button/link (though stopPropagation handles this usually)
    e.preventDefault();

    this.pos = {
      left: ele.scrollLeft,
      top: ele.scrollTop,
      x: e.clientX,
      y: e.clientY,
    };

    ele.style.cursor = 'grabbing';
    ele.style.userSelect = 'none';

    document.addEventListener('mousemove', this.onMouseMove);
    document.addEventListener('mouseup', this.onMouseUp);
  };

  private onMouseMove = (e: MouseEvent): void => {
    const ele = this.scrollAreaRef.current;
    if (!ele) return;
    const dx = e.clientX - this.pos.x;
    const dy = e.clientY - this.pos.y;
    ele.scrollTop = this.pos.top - dy;
    ele.scrollLeft = this.pos.left - dx;
  };

  private onMouseUp = (): void => {
    const ele = this.scrollAreaRef.current;
    if (ele) {
      ele.style.cursor = 'grab';
      ele.style.removeProperty('user-select');
    }
    document.removeEventListener('mousemove', this.onMouseMove);
    document.removeEventListener('mouseup', this.onMouseUp);
  };

  // --- ACTIONS ---

  private handleZoom = (delta: number): void => {
    this.setState(prev => ({ zoom: Math.max(0.2, prev.zoom + delta) }));
  };

  private handleResetZoom = (): void => {
    this.setState({ zoom: 1 }, () => setTimeout(this.handleCenterRoot, 100));
  };

  private handleCenterRoot = (): void => {
    if (this.rootCardRef.current && this.scrollAreaRef.current) {
        this.rootCardRef.current.scrollIntoView({ 
            behavior: 'smooth', block: 'center', inline: 'center' 
        });
    }
  }

  private handleFitToScreen = (): void => {
    if (this.containerRef.current && this.treeWrapperRef.current) {
      const containerW = this.containerRef.current.clientWidth;
      const contentW = this.treeWrapperRef.current.scrollWidth;
      const ratio = contentW > containerW ? (containerW - 40) / contentW : 1;
      this.setState({ zoom: ratio });
    }
  };

  private handlePrint = (): void => window.print();

  // --- RENDER ---

  private renderNodeContent = (node: IDepartmentNode, isRoot: boolean = false): React.ReactNode => {
    const accentColor = node.isContractorNode ? "#e3008c" : "#0083ca";
    const bgColor = node.isContractorNode ? "#fdf2f8" : "#faf9f8";
    const totalCount = node.employees.length + node.contractors.length;

    return (
      <div
        ref={isRoot ? this.rootCardRef : undefined}
        className={styles.nodeCard}
        style={{ borderTop: `6px solid ${accentColor}` }}
      >
        <div className={styles.header} style={{ backgroundColor: bgColor }}>
          <div className={styles.title}>{node.name}</div>
          <div className={styles.subTitle}>{totalCount} Members</div>
        </div>

        <div className={styles.body}>
          {node.employees.map((emp, i) => {
             const userEmail = emp.Email?.EMail || "";
             return (
              <div 
                key={emp.Id || i} 
                className={styles.empRow} 
                onClick={(e) => this.handleEmployeeClick(emp, e)}
                style={{ cursor: 'pointer', transition: 'background-color 0.2s' }}
                title="Click for details"
              >
                <img src={`/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`} />
                <div className={styles.empInfo}>
                  <div className={styles.empName}>{emp.Title}</div>
                  <div className={styles.empJob}>{emp.Job}</div>
                </div>
              </div>
             );
          })}
          {node.contractors.length > 0 && (
            <div className={styles.contractorSection}>
              <div className={styles.sectionHeader}>Saradnici ({node.contractors.length})</div>
              {node.contractors.map((emp, i) => {
                 const userEmail = emp.Email?.EMail || "";
                 return (
                  <div 
                    key={emp.Id || i} 
                    className={styles.empRow}
                    onClick={(e) => this.handleEmployeeClick(emp, e)}
                    style={{ cursor: 'pointer', transition: 'background-color 0.2s' }}
                    title="Click for details"
                  >
                    <img src={`/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`} style={{filter:'grayscale(100%)'}}/>
                    <div className={styles.empInfo}>
                      <div className={styles.empName}>{emp.Title}</div>
                      <div className={styles.empJob}>{emp.Job}</div>
                    </div>
                  </div>
                 );
              })}
            </div>
          )}
        </div>
      </div>
    );
  };

  private renderTreeNodes = (node: IDepartmentNode, visited: string[] = []): React.ReactNode => {
    if (visited.indexOf(node.id) !== -1) return null;
    const newVisited = [...visited, node.id];
    return (
      <TreeNode label={this.renderNodeContent(node)}>
        {node.children.map(child => this.renderTreeNodes(child, newVisited))}
      </TreeNode>
    );
  };

  public render(): React.ReactElement<IOrgTreeProps> {
    const { webPartWidth, transparentBackground } = this.props;
    const widthVal = webPartWidth ? `${webPartWidth}%` : '100%';
    const { selectedEmployee } = this.state;

    if (this.state.isLoading) return <div>Loading...</div>;
    if (this.state.error) return <div style={{color:'red'}}>{this.state.error}</div>;
    if (!this.state.rootNode) return <div style={{color:'red'}}>No Data Found</div>;

    const printCSS = `
      @media print {
        @page { size: landscape; margin: 0; }
        body * { visibility: hidden; }
        #org-tree-content, #org-tree-content * { visibility: visible; }
        #org-tree-content {
          position: fixed; left: 0; top: 0; width: 100%;
          transform: scale(0.6); transform-origin: top left;
        }
        .org-tree-toolbar { display: none !important; }
      }
    `;

    return (
      <div 
        ref={this.containerRef}
        className={styles.orgContainer}
        style={{ 
            width: widthVal, 
            background: transparentBackground ? 'transparent' : '#f3f2f1',
            border: transparentBackground ? 'none' : '1px solid #eaeaea'
        }}
      >
        <style>{printCSS}</style>

        <div className={`${styles.toolbar} org-tree-toolbar`}>
          <button onClick={() => this.handleZoom(0.1)} title="Zoom In">+</button>
          <button onClick={() => this.handleZoom(-0.1)} title="Zoom Out">-</button>
          <button onClick={this.handleResetZoom}>1:1</button>
          <button onClick={this.handleFitToScreen}>Fit</button>
          <button onClick={this.handlePrint}>🖨️</button>
        </div>

        {/* DRAG AREA */}
        <div 
          ref={this.scrollAreaRef}
          className={styles.scrollArea}
          onMouseDown={this.onMouseDown} 
        >
          <div 
            id="org-tree-content"
            ref={this.treeWrapperRef}
            className={styles.treeWrapper}
            style={{ transform: `scale(${this.state.zoom})` }}
          >
            <Tree
              lineWidth={'2px'}
              lineColor={'#0083ca'}
              lineBorderRadius={'10px'}
              label={this.renderNodeContent(this.state.rootNode, true)}
            >
              {this.state.rootNode.children.map(child => this.renderTreeNodes(child))}
            </Tree>
          </div>
        </div>

        {/* EMPLOYEE CARD POPUP */}
        <Dialog
          hidden={!selectedEmployee}
          onDismiss={this.closeDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Employee Details',
          }}
          modalProps={{
            isBlocking: false,
            dragOptions: { moveMenuItemText: 'Move', closeMenuItemText: 'Close', menu: undefined as any },
            styles: { main: { maxWidth: 450 } } // Limit total width
          }}
          minWidth={400}
        >
          {selectedEmployee && (
             <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '20px', width: '100%', overflow: 'hidden' }}>
                
                {/* Persona with Text Handling */}
                <div style={{ width: '100%', display: 'flex', justifyContent: 'center' }}>
                  <Persona
                    imageUrl={`/_layouts/15/userphoto.aspx?size=L&accountname=${selectedEmployee.Email?.EMail || ""}`}
                    text={selectedEmployee.Title}
                    secondaryText={selectedEmployee.Job}
                    tertiaryText={selectedEmployee.Dept}
                    size={PersonaSize.size100}
                    styles={{ 
                      primaryText: { fontWeight: 'bold' },
                      secondaryText: { whiteSpace: 'normal', height: 'auto', textAlign: 'center' }, // Allow wrap
                      tertiaryText: { whiteSpace: 'normal', height: 'auto', textAlign: 'center', marginTop: 4 }, // Allow wrap
                      details: { width: '100%', alignItems: 'center' } // Center align text
                    }}
                  />
                </div>
                
                <div style={{ width: '100%', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                   {selectedEmployee.Email?.EMail && (
                     <a href={`mailto:${selectedEmployee.Email.EMail}`} title={selectedEmployee.Email.EMail} style={{ textDecoration: 'none', color: '#333', display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', border: '1px solid #edebe9', borderRadius: '6px', transition: 'background 0.2s', backgroundColor: '#faf9f8' }}>
                        <Icon iconName="Mail" style={{ color: '#0083ca', fontSize: '20px', flexShrink: 0 }} />
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{selectedEmployee.Email.EMail}</span>
                     </a>
                   )}
                   
                   {selectedEmployee.Mobile && (
                     <a href={`tel:${selectedEmployee.Mobile}`} title={selectedEmployee.Mobile} style={{ textDecoration: 'none', color: '#333', display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', border: '1px solid #edebe9', borderRadius: '6px', transition: 'background 0.2s', backgroundColor: '#faf9f8' }}>
                        <Icon iconName="Phone" style={{ color: '#0083ca', fontSize: '20px', flexShrink: 0 }} />
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{selectedEmployee.Mobile}</span>
                     </a>
                   )}

                   {selectedEmployee.Email?.EMail && (
                     <a href={`https://teams.microsoft.com/l/chat/0/0?users=${selectedEmployee.Email.EMail}`} target="_blank" style={{ textDecoration: 'none', color: '#333', display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', border: '1px solid #edebe9', borderRadius: '6px', transition: 'background 0.2s', backgroundColor: '#faf9f8' }}>
                        <Icon iconName="TeamsLogo" style={{ color: '#464775', fontSize: '20px', flexShrink: 0 }} />
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>Chat in Teams</span>
                     </a>
                   )}
                </div>
             </div>
          )}
          <DialogFooter>
            <PrimaryButton onClick={this.closeDialog} text="Close" />
          </DialogFooter>
        </Dialog>

      </div>
    );
  }
}