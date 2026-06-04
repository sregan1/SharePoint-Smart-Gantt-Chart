declare module '*.module.scss' {
  const styles: { readonly [className: string]: string };
  export default styles;
}
