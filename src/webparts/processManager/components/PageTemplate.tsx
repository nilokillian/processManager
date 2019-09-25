import * as React from "react";
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps
} from "office-ui-fabric-react/lib/DocumentCard";
import { ImageFit } from "office-ui-fabric-react/lib/Image";

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
  constructor(props: IPageTemplateProps) {
    super(props);
  }

  // private _onCheckboxChange = (
  //   ev: React.FormEvent<HTMLElement>,
  //   isSelected: boolean
  // ) => {
  //   const { pageTitle } = this.props;

  //   // if (isChecked) {
  //   //   selectedTemplate.name = pageTemplate.name;
  //   //   selectedTemplate.title = pageTemplate.title;
  //   //   selectedTemplate.serverRelativeUrl = pageTemplate.serverRelativeUrl;

  //   //   onTemplateSelect(selectedTemplate);
  //   // } else {
  //   //   onTemplateSelect(selectedTemplate);
  //   // }
  //   onTemplateSelect(pageTemplate, isSelected);
  // };

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
          width: 300,
          height: 340
        }
      ]
    };

    return (
      <div>
        <DocumentCard>
          <DocumentCardPreview {...previewProps} />
          <DocumentCardTitle title="Policy" shouldTruncate={true} />
          <DocumentCardActivity
            activity="Created a few minutes ago"
            people={[{ name: pageTitle, profileImageSrc: templateImage.icon }]}
          />
        </DocumentCard>
      </div>
    );
  }
}
