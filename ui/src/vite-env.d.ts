/// <reference types="vite/client" />

interface OnlyOfficeEditorInstance {
	destroyEditor?: () => void;
}

interface OnlyOfficeDocEditorConstructor {
	new (placeholderId: string, config: Record<string, unknown>): OnlyOfficeEditorInstance;
	version?: () => string;
}

interface Window {
	DocsAPI?: {
		DocEditor: OnlyOfficeDocEditorConstructor;
	};
}
