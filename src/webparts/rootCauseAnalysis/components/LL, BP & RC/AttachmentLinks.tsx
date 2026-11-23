/*eslint-disable*/
import * as React from 'react';

export interface IAttachmentLink {
  FileName: string;
  ServerRelativeUrl: string;
}

interface IAttachmentLinksProps {
  attachments?: IAttachmentLink[];
}

const AttachmentLinks: React.FC<IAttachmentLinksProps> = ({ attachments }) => {
  if (!attachments || attachments.length === 0) {
    return <span>-</span>;
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
      {attachments.map((file, index) => (
        file.ServerRelativeUrl ? (
          <a
            key={`${file.FileName}-${index}`}
            href={file.ServerRelativeUrl}
            target="_blank"
            rel="noopener noreferrer"
          >
            {file.FileName}
          </a>
        ) : (
          <span key={`${file.FileName}-${index}`}>{file.FileName}</span>
        )
      ))}
    </div>
  );
};

export default AttachmentLinks;
