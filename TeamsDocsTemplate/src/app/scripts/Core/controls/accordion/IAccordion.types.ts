
export interface IAccordionProps {
  defaultCollapsed?: boolean;
  title: string;
  className?: string;
  onExpand: () => any;
  id: string;
  selected?: boolean;
  childCount?: number;
  foldertype?: string;
}


export interface IAccordionState {
  expanded: boolean;
  foldertype?: string;
  iconName?: string;
  // ChildCount: number;
}


