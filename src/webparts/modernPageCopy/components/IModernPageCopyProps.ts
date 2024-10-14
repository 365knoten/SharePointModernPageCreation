export interface IModernPageCopyProps {
  copyPage:(name:string)=>Promise<void>
  fieldTitle: string;
}
