import * as React from "react";
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
  DocumentCardDetails
} from "office-ui-fabric-react/lib/DocumentCard";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import { DefaultButton } from "office-ui-fabric-react";

const templateImage: ITemplateImage = {
  image: require("../../images/templateImg.png"),
  icon: require("../../images/templateIcon.jpg")
};

export interface ITemplate {
  name: string;
  title: string;
  serverRelativeUrl: string;
}

export interface ITemplateImage {
  image: string;
  icon: string;
}
export interface IPageTemplateProps {
  pageTitle: string;
  url: string;
}

export default class PageTemplate extends React.Component<
  IPageTemplateProps,
  {}
> {
  private _openNewTab = (url: string) => {
    window.open(url, "_blank");
  };

  constructor(props: IPageTemplateProps) {
    super(props);
  }

  public render(): JSX.Element {
    const { pageTitle, url } = this.props;
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          name: `${pageTitle}.aspx`,
          linkProps: {
            href: url,
            target: "_blank"
          },
          previewImageSrc: templateImage.image,
          imageFit: ImageFit.cover,
          width: 320,
          height: 240
        }
      ]
    };

    return (
      <DocumentCard>
        <DocumentCardPreview {...previewProps} />
        <DocumentCardDetails>
          <DocumentCardActivity
            activity={previewProps.previewImages[0].linkProps.href}
            people={[{ name: pageTitle, profileImageSrc: templateImage.icon }]}
          />
          <DefaultButton text="Edit" onClick={() => this._openNewTab(url)} />
        </DocumentCardDetails>
      </DocumentCard>
    );
  }
}
