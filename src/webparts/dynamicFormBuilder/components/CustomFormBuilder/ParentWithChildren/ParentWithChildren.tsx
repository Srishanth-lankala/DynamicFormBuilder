import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./ParentWithChildren.module.scss";

interface ChildItem {
  id: string | number;
  name: string;
  ChildFormJson: any[];
  AppCode : string;
  ChildOrder : string | number;
}

interface ParentWithChildrenProps {
  parentName: string;
  children: ChildItem[];
  onChildClick?: (child: ChildItem) => void;
  // AppCode: string;
}

const ParentWithChildren: React.FC<ParentWithChildrenProps> = ({
  parentName,
  children,
  onChildClick,
}) => {
  const [activeChild, setActiveChild] = useState<string | number | null>(null);
  const [isOpen, setIsOpen] = useState<boolean>(true); // default open;

  useEffect(() => { 
    if (children.length > 0 && activeChild === null) {
      setActiveChild(children[0].id);
      if (onChildClick) {
        onChildClick(children[0]);
      }
    }
  }, [children, activeChild, onChildClick]);

  return (
    <div className={styles.parentContainer}>
      {/* Header with toggle */}
      <div
        className={styles.header}
        onClick={() => setIsOpen((prev) => !prev)}
      >
        {parentName}
        <span className={styles.toggleBtn}>
          {isOpen ? "Hide ▲" : "Show ▼"}
        </span>
      </div>

      {/* Sliding panel */}
      <div
        className={styles.panel}
        style={{ maxHeight: isOpen ? "500px" : "0px" }}
      >
        {children.length > 0 ? (
          <div className={styles.childTabsWrapper}>
            {children.map((child) => (
              <div
                key={child.id}
                className={`${styles.childTab} ${
                  activeChild === child.id ? styles.activeTab : ""
                }`}
                onClick={() => {
                  setActiveChild(child.id);
                  if (onChildClick) onChildClick(child);
                }}
              >
                {child.name}
              </div>
            ))}
          </div>
        ) : (
          <div className={styles.noChildren}>No child forms available</div>
        )}
      </div>
    </div>
  );
};

export default ParentWithChildren;
