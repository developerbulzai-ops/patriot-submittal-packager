export interface LineItem {
  description: string;
  /** Page number in the OUTPUT PDF (Patriot cover = 1, so supplier pages shift +1) */
  startPage: number;
  endPage: number;
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
  category: string;
  lineItems: LineItem[];
}
