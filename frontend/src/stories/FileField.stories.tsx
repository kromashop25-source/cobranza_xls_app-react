import { useState, type ComponentProps } from "react";
import type { Meta, StoryObj } from "@storybook/react-vite";

import FileField from "../components/FileField";

const meta = {
  title: "Forms/FileField",
  component: FileField,
  tags: ["autodocs"],
  args: {
    label: "Archivo XLS",
    accept: ".xls",
    hint: "Selecciona un archivo de prueba",
  },
} satisfies Meta<typeof FileField>;

export default meta;

type Story = StoryObj<typeof meta>;

type FileFieldProps = ComponentProps<typeof FileField>;

function FileFieldPreview(args: FileFieldProps) {
  const [file, setFile] = useState<File | null>(null);
  return <FileField {...args} value={file} onChange={setFile} />;
}

export const Default: Story = {
  args: {
    value: null,
    onChange: () => {
      // handled inside preview component state
    },
  },
  render: (args) => <FileFieldPreview {...args} />,
};
