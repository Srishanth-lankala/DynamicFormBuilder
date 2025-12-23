import React, { useEffect, useRef, useState } from 'react';
import './SideBar.css';

import HomeIcon from '../../../assets/Images/SidebarIcons/HomeIcon.svg';
import FormIcon from '../../../assets/Images/SidebarIcons/FormIcon.svg';
import WorkflowIcon from '../../../assets/Images/SidebarIcons/WorkflowIcon.svg';
import ApprovalsIcon from '../../../assets/Images/SidebarIcons/ApprovalsIcon.svg';
import RequestsIcon from '../../../assets/Images/SidebarIcons/RequestsIcon.svg';
import FormDashBoardIcon from '../../../assets/Images/SidebarIcons/FormDashBoardIcon.svg';
import ReportsIcon from '../../../assets/Images/SidebarIcons/ReportsIcon.svg';
import MyZoneIcon from '../../../assets/Images/MyZoneIcon.svg';
import SingleForm from '../../../assets/Images/SidebarIcons/SingleForm.svg';
import MultiForm from '../../../assets/Images/SidebarIcons/MultiForm.svg';
import InvoiceIcon from '../../../assets/Images/SidebarIcons/InvoiceIcon.svg';
import { useNavigate } from 'react-router-dom';
import { sp } from '@pnp/sp';
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
import { Toast } from 'primereact/toast';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faBars, faChevronDown, faChevronUp } from '@fortawesome/free-solid-svg-icons';
import { fetchFormCount } from '../CustomHooks/TransactionCount';

interface SidebarItem {
  Menu: string;
  Icon?: string;
  path?: string;
  hasNotification?: boolean;
  children?: SidebarItem[];
}

interface SidebarProps {
  activeMenu: string | null;
  onItemClick?: (menu: string) => void;
}

const SideBar: React.FC<SidebarProps> = ({ activeMenu }) => {
  const [redDot, setRedDot] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  // const [createFormToasterr, setCreateFormToasterr] = useState<boolean>(false);
  const [isOpen, setIsOpen] = useState(window.innerWidth >= 1024);
  const [isTabView, setIsTabView] = useState(window.innerWidth <= 768);
  const [Sidebaritems, setSidebaritems] = useState<SidebarItem[]>([]);
  const [openSubMenu, setOpenSubMenu] = useState<string | null>(null);
  const navigate = useNavigate();
  const toast = useRef<Toast>(null);

  const basesidebarItems: SidebarItem[] = [
    { Menu: 'Home', Icon: HomeIcon, path: '/LoadForm' },
    {
      Menu: 'Create Form',
      Icon: FormIcon,
      children: [
        { Menu: 'Single Form', Icon: SingleForm, path: '/CreateForm' },
        { Menu: 'Multi Form', Icon: MultiForm, path: '/CreateForm' },
      ],
    },
    { Menu: 'Create Workflow', Icon: WorkflowIcon, path: '/ConfigureWorkflow' },
    { Menu: 'Pending Approvals', Icon: ApprovalsIcon, path: '/InProcess', hasNotification: redDot },
    { Menu: 'My Requests', Icon: RequestsIcon, path: '/Draft' },
    { Menu: 'Form DashBoard', Icon: FormDashBoardIcon, path: '/FormDashBoard' },
    { Menu: 'Generate Reports', Icon: ReportsIcon, path: '/Reports' },
    { Menu: 'PDF Template', Icon: InvoiceIcon, path: '/TemplateMapper' },
  ];

  const EditFormData = {
    isEdit: false,
    AppCode: '',
    AppName: '',
    FormType: '',
    FormJSON: '',
    Description: '',
  };

  // Fetch workflow items to set red dot
  const fetchData = async () => {
    try {
      const currentUser = await sp.web.currentUser.get();
      const currentUserGroups = await sp.web.currentUser.groups.get();
      const Domain = await fetchTenantUser();
      const allItems = await sp.web.lists
        .getByTitle('WorkFlowProcessData')
        .items.orderBy('Modified', false)
        .select('ID', 'Title', 'Status', 'Created', 'Modified', 'Author/Title', 'CurApprover/Title', 'CurApproverId', 'AppName')
        .expand('Author', 'CurApprover')
        .filter(`Domain eq '${Domain.TenantUsers}' and (Status ne 'Draft' and Status ne 'Cancelled' and Status ne 'Returned' and Status ne 'Submitted' and Status ne 'Rejected' and Status ne 'Completed') `)
        .getAll();
      let hasPendingItems = false;
      hasPendingItems = allItems.some(item => item.CurApproverId === currentUser.Id);
      if (!hasPendingItems && currentUserGroups?.length > 0) {
        const userGroupIds = new Set(currentUserGroups.map(g => g.Id));
        hasPendingItems = allItems.some(
          item => item.CurApproverId && userGroupIds.has(item.CurApproverId)
        );
      }
      setRedDot(hasPendingItems);
    } catch (error) {
      console.error('Error fetching SharePoint data:', error);
    }
  };

  // Check admin rights
  const isAdminUser = async () => {
    const currentUser = await sp.web.currentUser.get();
    const authUsers = await sp.web.lists
      .getByTitle('AuthList')
      .items.filter(`AdminUser eq 'Yes' and AuthName/Id eq ${currentUser.Id}`)
      .select('AuthName/Id')
      .expand('AuthName')
      .getAll();
    setIsAdmin(authUsers.length > 0);
  };

  // Check license-based form limit


  useEffect(() => {
    const getSidebarItems = async () => {
      const domain = await fetchTenantUser();
      const items = [...basesidebarItems];
      if (domain.TenantUsers === 'cloudangles.com') {
        items.splice(1, 0, { Menu: 'My Zone', Icon: MyZoneIcon, path: '/MyZone' });
      }
      setSidebaritems(items);
    };
    void getSidebarItems();
  }, [redDot]);

  useEffect(() => {
    const handleResize = () => {
      const width = window.innerWidth;
      setIsTabView(width <= 768);
      setIsOpen(width >= 1024);
    };
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    void fetchData();
    void isAdminUser();

  }, []);

  // Handle Create Form subscription check
  const handleCreateFormWorkflow = async () => {
    const tenant = await fetchTenantUser();
    const domain = tenant.TenantUsers;
    const { formcount, limitReached } = await fetchFormCount(domain);
    if (!limitReached) {
      sessionStorage.setItem('EditFormData', JSON.stringify(EditFormData));
      return true;
    } else {
      toast.current?.clear();
      toast.current?.show({
        severity: 'warn',
        summary: 'Freemium subscription limit reached',
        detail: `Form creation limit ${formcount} exceeded. Please upgrade your license.`,
        life: 3000,
        style: { height: 90 }
      });
      return false;
    }
  };

  // Sidebar item click logic
  const sidebarclick = async (item: SidebarItem) => {
    if (!isAdmin && (item.Menu === 'Create Form' || item.Menu === 'Create Workflow')) return;

    if (item.Menu === 'Create Form') {
      setOpenSubMenu((prev) => (prev === 'Create Form' ? null : 'Create Form'));
      return;
    }

    if (item.Menu === 'Single Form') {
      const canProceed = await handleCreateFormWorkflow();
      try {
        sessionStorage.removeItem('FormMasterData');
        sessionStorage.setItem('isMultiForm', 'false');
      } catch (error) {
        console.error('Error while setting multi flag:', error);
      }
      if (!canProceed) return;
      if (item.path) navigate(item.path);
      return;
    }

    if (item.Menu === 'Multi Form') {
      const canProceed = await handleCreateFormWorkflow();
      if (!canProceed) return;

      try {
        const tenantt = await fetchTenantUser();

        const formMasterItems = await sp.web.lists
          .getByTitle('FormMaster')
          .items.select('ID,AppCode,AppName,FormJSON,IsParentForm')
          // .filter(`ChildOrder eq null and Domain eq '${tenantt.TenantUsers}'`)
          .filter(`ChildOrder eq null and IsParentForm ne 1 and Domain eq '${tenantt.TenantUsers}' and VisibilityFlag ne 0 and isWorkflowRequired ne 'No'`)
          .getAll();

        sessionStorage.setItem("FormMasterData", JSON.stringify(formMasterItems));
        sessionStorage.setItem("isMultiForm", "true");
        sessionStorage.setItem("isNewForm", "true");
        //console.log("Multi Form: Set isMultiForm = true with", formMasterItems.length, "forms");
      } catch (error) {
        console.error("Error fetching MultiForm data:", error);
      }

      if (item.path) navigate(item.path);
      return;
    }

    if (isTabView) setIsOpen(false);
    if (item.path) navigate(item.path);
  };

  return (
    <div className={`sidebar ${isOpen ? 'expanded' : 'collapsed'}`}>
      <Toast ref={toast} />

      {/* Toggle button (only hamburger for smaller view) */}
      {isTabView && (
        <div className="sidebar-toggle">
          <FontAwesomeIcon
            icon={faBars}
            className="text-2xl cursor-pointer"
            onClick={() => setIsOpen(!isOpen)}
          />
        </div>
      )}

      {Sidebaritems.map((item, index) => (
        <div key={index}>
          <div
            className={`sidebar-item ${activeMenu === item.Menu ? 'active' : ''}`}
            style={{
              ...(!isAdmin && (item.Menu === 'Create Form' || item.Menu === 'Create Workflow')
                ? { backgroundColor: '#212B467D', cursor: 'not-allowed', opacity: 0.5 }
                : { cursor: 'pointer' }),
            }}
            onClick={() => sidebarclick(item)}
          >
            {item.Icon && <img src={item.Icon} alt={item.Menu} className="menu-icon" />}
            {isOpen && <span>{item.Menu}</span>}
            {item.Menu === 'Create Form' && isOpen && (
              <FontAwesomeIcon
                icon={openSubMenu === 'Create Form' ? faChevronUp : faChevronDown}
                className="ml-auto"
              />
            )}
            {item.hasNotification && <div className="red-dot" />}
          </div>

          {/* Submenu */}
          {item.children && openSubMenu === item.Menu && (
            <div className="submenu">
              {item.children.map((child, idx) => (
                <div
                  key={idx}
                  className={`submenu-item`}
                  style={{ cursor: 'pointer' }}
                  onClick={() => sidebarclick(child)}
                >
                  {child.Icon && <img src={child.Icon} alt={child.Menu} className="menu-icon" />}
                  {isOpen && <span>{child.Menu}</span>}
                </div>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
};

export default SideBar;