/** Re-export — Espárrago MP usa el análisis compartido. */
export {
  buildCartillaAnalysis as buildEsparragoMpCartillaAnalysis,
  htmlCartillaAnalysisModal,
  htmlCartillaAnalysisPanel,
  createCartillaAnalysisController,
  deriveFilasConErrorFromDom,
  filterFilasConErrorExcludingSapOnly,
  headersToAnalysisColumns
} from "../shared/cartilla-analysis.js";
