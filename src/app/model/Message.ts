import {EmailAddress} from './emailAddress';

export class Message {
  message:         MessageClass;
  saveToSentItems: string;
}

export class MessageClass {
  subject:      string;
  body:         Body;
  toRecipients: ToRecipient[];
}

export class Body {
  contentType: string;
  content:     string;
}

export class ToRecipient {
  emailAddress: EmailAddress;
}


