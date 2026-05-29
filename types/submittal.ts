export interface LineItem {
  description: string;
  startPage: number; // page in the final assembled document (title=1, blank=2, data starts at 3)
  endPage: number;
  warning?: string;  // set when AI detects non-consecutive pages for this product
}

export interface CategoryGroup {
  name: string;
  lineItems: LineItem[];
}

export interface SubmittalData {
  jobNo: string;
  date: string;
  recipient: {
    company: string;
    attention: string;
    address1: string;
    city: string;
  };
  subject: {
    projectName: string;
    location: string;
  };
  categories: CategoryGroup[];
}
