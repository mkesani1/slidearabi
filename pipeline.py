import logging
import time
from pathlib import Path
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Callable, Any
from pptx import Presentation

logger = logging.getLogger(__name__)

# These modules are built by other agents and will be available at runtime
# Import them at the top but handle ImportError gracefully for testing
try:
    from slidearabi.property_resolver import PropertyResolver
    from slidearabi.layout_analyzer import LayoutAnalyzer
    from slidearabi.template_registry import TemplateRegistry
    from slidearabi.rtl_transforms import MasterLayoutTransformer, SlideContentTransformer
    from slidearabi.typography import TypographyNormalizer
    from slidearabi.structural_validator import StructuralValidator
    from slidearabi.models import (
        ResolvedPresentation, TransformReport, ValidationReport,
        PipelineConfig, PipelineResult
    )
except ImportError as e:
    logger.warning(f"Some v2 modules not yet available: {e}")
    
    # Stub models for development/testing if real ones aren't available
    @dataclass
    class PipelineConfig:
        input_path: str
        output_path: str
        translate_fn: Optional[Callable[[List[str]], Dict[str, str]]] = None
        skip_translation: bool = False
        max_font_reduction_pct: float = 20.0
        log_level: str = 'INFO'
        enable_telemetry: bool = False
        
    @dataclass
    class PipelineResult:
        success: bool
        output_path: Optional[str]
        phase_reports: Dict[str, Any]
        validation_report: Optional[Any]
        total_duration_ms: float
        error: Optional[str] = None

# Phase 6 (VQA) and LLM translation — optional but wired when available
try:
    from slidearabi.visual_qa import run_vqa, VQAReport
    HAS_VQA = True
except ImportError:
    HAS_VQA = False

try:
    from slidearabi.llm_translator import DualLLMTranslator, TranslatorConfig
    HAS_LLM_TRANSLATOR = True
except ImportError:
    HAS_LLM_TRANSLATOR = False


class SlideArabiPipeline:
    """
    Template-first, deterministic RTL transformation pipeline.
    
    Phases execute as a DAG — each phase runs exactly once and produces
    immutable output consumed by subsequent phases. No fix loops.
    
    Phase 0: Parse & Resolve → ResolvedPresentation
    Phase 1: Translate (AI — GPT-5.2 primary, Claude 4.6 QA) → translation_map
    Phase 2: Master & Layout Transformation (deterministic) → TransformedPresentation
    Phase 3: Slide Content Transformation (deterministic) → TransformedPresentation
    Phase 4: Typography Normalization (deterministic) → FinalPresentation
    Phase 5: Structural Validation (read-only) → ValidationReport
    Phase 6: Visual QA (Gemini 3.1 Pro + Claude 4.6 dual-pass) → VQAReport
    """
    
    def __init__(self, config: 'PipelineConfig'):
        self.config = config
        self._phase_reports = {}
        
        # Configure logging
        numeric_level = getattr(logging, config.log_level.upper(), logging.INFO)
        logger.setLevel(numeric_level)
        
    def run(self) -> 'PipelineResult':
        """
        Execute the full pipeline. Returns PipelineResult.
        
        This is synchronous. Each phase completes before the next begins.
        Phases are NOT async because there's no I/O wait between them
        (except Phase 1 translation which calls an external API).
        """
        pipeline_start = time.monotonic()
        logger.info(f"Starting SlideArabi pipeline for {self.config.input_path}")
        
        try:
            # 1. Load presentation
            logger.info("Loading presentation...")
            try:
                prs = Presentation(self.config.input_path)
            except Exception as e:
                raise ValueError(f"Failed to load presentation: {e}")
            
            # Phase 0: Parse & Resolve
            resolved_prs = self._phase_0_resolve(prs)
            
            # Phase 1: Translate
            translation_map = self._phase_1_translate(resolved_prs)
            
            # Phase 2: Transform Masters & Layouts
            p2_report = self._phase_2_transform_masters_layouts(prs, resolved_prs)
            
            # Phase 3: Transform Slide Content
            p3_report = self._phase_3_transform_slides(prs, resolved_prs, translation_map)
            
            # Phase 4: Typography Normalization
            p4_report = self._phase_4_typography(prs)
            
            # Phase 5: Validate
            val_report = self._phase_5_validate(prs, resolved_prs)
            
            # Save presentation
            logger.info(f"Saving transformed presentation to {self.config.output_path}...")
            try:
                # Ensure directory exists
                Path(self.config.output_path).parent.mkdir(parents=True, exist_ok=True)
                prs.save(self.config.output_path)
            except Exception as e:
                raise IOError(f"Failed to save presentation: {e}")
            
            # Phase 6: Visual QA (dual-pass: Gemini + Claude)
            vqa_report = self._phase_6_vqa()
            
            total_duration = (time.monotonic() - pipeline_start) * 1000
            logger.info(f"Pipeline completed successfully in {total_duration:.0f}ms")
            
            return PipelineResult(
                success=True,
                output_path=self.config.output_path,
                phase_reports=self._phase_reports,
                validation_report=val_report,
                total_duration_ms=total_duration
            )
            
        except Exception as e:
            logger.error(f"Pipeline failed: {e}", exc_info=True)
            total_duration = (time.monotonic() - pipeline_start) * 1000
            
            return PipelineResult(
                success=False,
                output_path=None,
                phase_reports=self._phase_reports,
                validation_report=None,
                total_duration_ms=total_duration,
                error=str(e)
            )
        
    def _phase_0_resolve(self, prs: Presentation) -> Any:
        """Phase 0: Parse the presentation and resolve all inherited properties."""
        start_time = time.monotonic()
        logger.info("Phase 0: Resolving properties...")
        
        try:
            resolver = PropertyResolver(prs)
            resolved_prs = resolver.resolve_presentation()
            
            # Optionally add layout analysis here if needed for phase 0 output
            analyzer = LayoutAnalyzer()
            # analyzer could mutate or wrap resolved_prs
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_0_resolve', duration, {"status": "success", "slides_resolved": len(prs.slides)})
            return resolved_prs
            
        except NameError as e:
            # Fallback if modules aren't available for tests
            logger.warning(f"Phase 0 stubbed: {e}")
            self._log_phase('phase_0_resolve', 0, {"status": "stubbed"})
            return None
            
    def _phase_1_translate(self, resolved: Any) -> Dict[str, str]:
        """Phase 1: Extract text and call translation function."""
        start_time = time.monotonic()
        logger.info("Phase 1: Translation...")
        
        if self.config.skip_translation:
            logger.info("Translation skipped per config.")
            self._log_phase('phase_1_translate', 0, {"status": "skipped"})
            return {}
            
        if not self.config.translate_fn:
            logger.warning("No translation function provided.")
            self._log_phase('phase_1_translate', 0, {"status": "no_function"})
            return {}
            
        texts_to_translate = self._extract_texts(resolved)
        logger.info(f"Extracted {len(texts_to_translate)} strings for translation.")
        
        try:
            translation_map = self.config.translate_fn(texts_to_translate)
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_1_translate', duration, {
                "status": "success", 
                "strings_translated": len(translation_map)
            })
            return translation_map
            
        except Exception as e:
            logger.error(f"Translation failed: {e}")
            raise RuntimeError(f"Phase 1 (Translation) failed: {e}")
        
    def _phase_2_transform_masters_layouts(self, prs: Presentation, resolved: Any) -> Any:
        """Phase 2: Transform masters and layouts deterministically."""
        start_time = time.monotonic()
        logger.info("Phase 2: Transforming masters and layouts...")
        
        try:
            registry = TemplateRegistry()
            transformer = MasterLayoutTransformer(prs, registry)
            
            transformer.transform_all_masters()
            report = transformer.transform_all_layouts()
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_2_transform_masters', duration, {"status": "success"})
            return report
            
        except NameError as e:
            logger.warning(f"Phase 2 stubbed: {e}")
            self._log_phase('phase_2_transform_masters', 0, {"status": "stubbed"})
            return None
        
    def _phase_3_transform_slides(self, prs: Presentation, resolved: Any, translations: Dict[str, str]) -> Any:
        """Phase 3: Transform slide content deterministically."""
        start_time = time.monotonic()
        logger.info("Phase 3: Transforming slide content...")
        
        try:
            # Layout analyzer might be needed to get classifications
            analyzer = LayoutAnalyzer()
            layout_classifications = analyzer.classify_slides(prs)
            
            transformer = SlideContentTransformer(prs, layout_classifications, translations)
            report = transformer.transform_all_slides()
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_3_transform_slides', duration, {"status": "success"})
            return report
            
        except NameError as e:
            logger.warning(f"Phase 3 stubbed: {e}")
            self._log_phase('phase_3_transform_slides', 0, {"status": "stubbed"})
            return None
        
    def _phase_4_typography(self, prs: Presentation) -> Any:
        """Phase 4: Normalize typography for Arabic."""
        start_time = time.monotonic()
        logger.info("Phase 4: Normalizing typography...")
        
        try:
            normalizer = TypographyNormalizer(prs, max_reduction_pct=self.config.max_font_reduction_pct)
            report = normalizer.normalize_all()
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_4_typography', duration, {"status": "success"})
            return report
            
        except NameError as e:
            logger.warning(f"Phase 4 stubbed: {e}")
            self._log_phase('phase_4_typography', 0, {"status": "stubbed"})
            return None
        
    def _phase_5_validate(self, prs: Presentation, resolved: Any) -> Any:
        """Phase 5: Read-only structural validation."""
        start_time = time.monotonic()
        logger.info("Phase 5: Structural validation...")
        
        try:
            validator = StructuralValidator(prs, resolved)
            report = validator.validate()
            
            duration = (time.monotonic() - start_time) * 1000
            
            log_data = {
                "status": "success",
                "passed": report.passed,
                "errors": report.errors,
                "warnings": report.warnings
            }
            self._log_phase('phase_5_validate', duration, log_data)
            
            if not report.passed:
                logger.warning(f"Validation failed with {report.errors} errors.")
                # We don't fail the pipeline, just report it
                
            return report
            
        except NameError as e:
            logger.warning(f"Phase 5 stubbed: {e}")
            self._log_phase('phase_5_validate', 0, {"status": "stubbed"})
            return None
        
    def _extract_texts(self, resolved: Any) -> List[str]:
        """Extract all translatable text strings from the resolved presentation."""
        texts = []
        if not resolved:
            return texts
            
        try:
            # Assuming resolved is a ResolvedPresentation with slides -> shapes -> paragraphs -> runs
            for slide in resolved.slides:
                for shape in slide.shapes:
                    for para in shape.paragraphs:
                        # Combine runs into a single translatable string per paragraph
                        para_text = "".join(run.text for run in para.runs if run.text).strip()
                        if para_text:
                            texts.append(para_text)
                            
            # Deduplicate while preserving order
            seen = set()
            unique_texts = []
            for t in texts:
                if t not in seen:
                    seen.add(t)
                    unique_texts.append(t)
            return unique_texts
            
        except AttributeError:
            # Fallback if resolved structure isn't exactly as expected
            logger.warning("Could not extract texts from resolved presentation structure")
            return []
        
    def _phase_6_vqa(
        self,
    ) -> Any:
        """Phase 6: Dual-pass Visual QA (Gemini 3.1 Pro + Claude Sonnet 4.6).
        
        Compares the original and converted PPTX visually:
        - Pass 1: Gemini evaluates slide composites in parallel
        - Pass 2: Claude adjudicates Gemini findings
        - Layer 3: Deterministic reconciliation
        - Remediates FAIL slides (font reduction, alignment, etc.)
        - Logs all issues to JSONL for root-cause analysis
        """
        start_time = time.monotonic()
        logger.info("Phase 6: Visual QA...")
        
        if not HAS_VQA:
            logger.warning("Phase 6 skipped: visual_qa module not available")
            self._log_phase('phase_6_vqa', 0, {"status": "module_unavailable"})
            return None
            
        if not self.config.input_path or not self.config.output_path:
            logger.warning("Phase 6 skipped: missing input/output paths")
            self._log_phase('phase_6_vqa', 0, {"status": "missing_paths"})
            return None
        
        # Set up issue log path alongside the output
        output_dir = Path(self.config.output_path).parent
        deck_stem = Path(self.config.output_path).stem
        issue_log_path = str(output_dir / f"{deck_stem}_vqa_issues.jsonl")
        
        try:
            report = run_vqa(
                original_pptx=self.config.input_path,
                converted_pptx=self.config.output_path,
                issue_log_path=issue_log_path,
                deck_name=deck_stem,
                enable_dual_pass=True,
            )
            
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_6_vqa', duration, {
                "status": "success",
                "overall_rating": report.overall_rating.value,
                "pass_rate": round(report.pass_rate, 1),
                "slides_reviewed": report.slides_reviewed,
                "fail_count": report.fail_count,
                "remediation_attempted": report.remediation_attempted,
                "remediation_successful": report.remediation_successful,
                "issues_logged": report.issues_logged,
            })
            
            logger.info(report.summary())
            return report
            
        except Exception as e:
            logger.warning(f"Phase 6 VQA failed (non-fatal): {e}")
            duration = (time.monotonic() - start_time) * 1000
            self._log_phase('phase_6_vqa', duration, {
                "status": "error",
                "error": str(e),
            })
            return None

    def _log_phase(self, phase_name: str, duration_ms: float, report: Any):
        """Log phase completion with timing and summary."""
        self._phase_reports[phase_name] = {
            "duration_ms": duration_ms,
            "report": report
        }
        logger.debug(f"{phase_name} completed in {duration_ms:.1f}ms: {report}")
