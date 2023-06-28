import { ITextFieldStyles } from "office-ui-fabric-react";
//import { IMeetingMinutesFormProps } from "./meetingMinutesForm/components/IMeetingMinutesFormProps";

export const LISTNAME ="Meeting Minutes";
export const METADATA_LISTNAME ="Metadata";
export const CONTACTS_LIST_NAME ="Contacts";
export const LIBRARYNAME = "MeetingMinutesSupportingFiles"//"Meeting Minutes Documents";
//export const SUBMIT_REDIRECT = (props: IMeetingMinutesFormProps): URL => {return new URL(props.context.pageContext.web.absoluteUrl +"/SitePages/Home.aspx");};
//export const CANCEL_REDIRECT = (props: IMeetingMinutesFormProps): URL => {return new URL(props.context.pageContext.web.absoluteUrl);};

export const CUSTOMER_LIST_ID ="acea17d5-8c92-4eec-80c4-289c6faa4cea";
export const CUSTOMER_URL = "https://k6931.sharepoint.com/sites/CommercialHub"
export const CUSTOMER_LISTNAME ="Customers";


export const modules = {  
    toolbar: [  
        [{  
            'header': [1, 2, 3, false]  
        }],  
        ['bold', 'italic', 'underline', 'strike', 'blockquote'],  
         
        [{  
            'list': 'ordered'  
        }, {  
            'list': 'bullet'  
        }, {  
            'indent': '-1'  
        }, {  
            'indent': '+1'  
        }],  
        ['image']  
        
    ],  
};
export const formats = ['header', 'bold', 'italic', 'underline', 'strike', 'blockquote', 'list', 'bullet', 'indent', 'image', 'background', 'color']; 

export const textFieldStyles: Partial<ITextFieldStyles> = {
    field: {
      width: '600px', // Adjust the desired width
    },
  };
  