import { html } from "lit";
import type { Meta, StoryObj } from "@storybook/web-components-vite";
import "./pptx-browser-demo";

type DemoArgs = {
  slideWidth: number;
  slideHeight: number;
  scaleToFit: boolean;
  letterbox: boolean;
  autoloadSample: boolean;
};

const meta: Meta<DemoArgs> = {
  title: "Library/PptxToHtml",
  tags: ["docsPage"],
  render: (args) => html`
    <pptx-browser-demo
      slide-width=${args.slideWidth}
      slide-height=${args.slideHeight}
      ?scale-to-fit=${args.scaleToFit}
      ?letterbox=${args.letterbox}
      ?autoload-sample=${args.autoloadSample}
    ></pptx-browser-demo>
  `,
  argTypes: {
    slideWidth: {
      control: { type: "number", min: 320, max: 1280, step: 40 },
    },
    slideHeight: {
      control: { type: "number", min: 180, max: 720, step: 20 },
    },
    scaleToFit: { control: "boolean" },
    letterbox: { control: "boolean" },
    autoloadSample: { control: "boolean" },
  },
  args: {
    slideWidth: 960,
    slideHeight: 540,
    scaleToFit: true,
    letterbox: false,
    autoloadSample: true,
  },
};

export default meta;

type Story = StoryObj<DemoArgs>;

export const GeneratedSample: Story = {
  name: "Generated sample",
};

export const UploadYourOwnDeck: Story = {
  name: "Upload your own deck",
  args: {
    autoloadSample: false,
    slideWidth: 880,
    slideHeight: 495,
    scaleToFit: true,
    letterbox: false,
  },
};
