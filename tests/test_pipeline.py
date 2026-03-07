import unittest
from unittest.mock import patch, MagicMock
from slidearabi.pipeline import SlideArabiPipeline, PipelineConfig, PipelineResult

class TestPipeline(unittest.TestCase):
    
    @patch('slidearabi.pipeline.Presentation')
    def test_pipeline_success_path(self, mock_presentation):
        # Setup mocks
        mock_prs = MagicMock()
        mock_presentation.return_value = mock_prs
        
        # Test config
        config = PipelineConfig(
            input_path="input.pptx",
            output_path="output.pptx",
            skip_translation=True
        )
        
        pipeline = SlideArabiPipeline(config)
        
        # Override phase methods to return dummy data instead of relying on missing imports
        pipeline._phase_0_resolve = MagicMock(return_value="resolved_prs")
        pipeline._phase_1_translate = MagicMock(return_value={"Hello": "مرحبا"})
        pipeline._phase_2_transform_masters_layouts = MagicMock(return_value="p2_report")
        pipeline._phase_3_transform_slides = MagicMock(return_value="p3_report")
        pipeline._phase_4_typography = MagicMock(return_value="p4_report")
        pipeline._phase_5_validate = MagicMock(return_value=MagicMock(passed=True))
        
        # Run
        result = pipeline.run()
        
        # Verify
        self.assertTrue(result.success)
        self.assertEqual(result.output_path, "output.pptx")
        
        # Verify DAG execution order
        pipeline._phase_0_resolve.assert_called_once_with(mock_prs)
        pipeline._phase_1_translate.assert_called_once_with("resolved_prs")
        pipeline._phase_2_transform_masters_layouts.assert_called_once_with(mock_prs, "resolved_prs")
        pipeline._phase_3_transform_slides.assert_called_once_with(mock_prs, "resolved_prs", {"Hello": "مرحبا"})
        pipeline._phase_4_typography.assert_called_once_with(mock_prs)
        pipeline._phase_5_validate.assert_called_once_with(mock_prs, "resolved_prs")
        
        # Verify save
        mock_prs.save.assert_called_once_with("output.pptx")

    @patch('slidearabi.pipeline.Presentation')
    def test_pipeline_failure_path(self, mock_presentation):
        mock_presentation.side_effect = Exception("Corrupt file")
        
        config = PipelineConfig(
            input_path="bad.pptx",
            output_path="output.pptx"
        )
        
        pipeline = SlideArabiPipeline(config)
        result = pipeline.run()
        
        self.assertFalse(result.success)
        self.assertIn("Corrupt file", result.error)

if __name__ == '__main__':
    unittest.main()
