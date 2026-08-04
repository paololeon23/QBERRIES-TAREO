const currentYear = new Date().getFullYear();

export const appConfig = {
  appYear: currentYear,
  appName: "AGV - MI",
  appVersion: "1.0.0",
  defaultLanguage: "es-PE",
  supportedLanguages: ["es-PE", "en-US", "fr-MA", "zh-CN"],
  cacheBustingVersion: "2026080402",
  defaultRoute: "#/inicio",
  brandLogoPath: "presentation/images/logo.png",
  brandMarkPath: "presentation/images/logo%20-%20copia.png",
  faviconPath: "presentation/images/icono.png"
};
