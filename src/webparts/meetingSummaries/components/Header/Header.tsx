import * as React from 'react';
import "./Header.css"
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHeaderProps {
  Title: string;
  context: WebPartContext;
}

export function Header({ Title, context }: IHeaderProps) {
  return (
    <div dir="rtl" className="EOHeader" style={{ marginBottom: '20px' }}>
      <div className="EOHeaderContainer" style={{ textAlign: 'center' }}>
        <span className="EOHeaderText" style={{ fontSize: '24px', fontWeight: 'bold' }}>
          {Title}{context.isServedFromLocalhost ? ' (localhost)' : ''}
        </span>
      </div>
      <div className="EOLogoContainer" style={{ textAlign: 'center' }}>
        {/* Logo goes here */}
      </div>
    </div>
  );
}
