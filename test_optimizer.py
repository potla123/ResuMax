import unittest
from unittest.mock import patch, MagicMock
import resume_optimizer

class TestResumeOptimizer(unittest.TestCase):
    
    @patch("resume_optimizer.requests.get")
    def test_scrape_job_description(self, mock_get):
        # Mocking requests.get to return fake html content
        mock_response = MagicMock()
        mock_response.text = """
        <html>
            <head><title>Job details</title></head>
            <body>
                <header>Header content</header>
                <nav>Navigation</nav>
                <main>
                    <h1>Senior Software Engineer</h1>
                    <p>We are looking for a senior engineer with Python and React experience.</p>
                    <p>Must have 5+ years of experience and knowledge of REST APIs.</p>
                </main>
                <footer>Footer stuff</footer>
            </body>
        </html>
        """
        mock_response.raise_for_status = MagicMock()
        mock_get.return_value = mock_response

        url = "https://example.com/job/1"
        result = resume_optimizer.scrape_job_description(url)
        
        # Verify cleaning: script/nav/header/footer tags should be removed
        self.assertIn("Senior Software Engineer", result)
        self.assertIn("Python and React", result)
        self.assertNotIn("Header content", result)
        self.assertNotIn("Navigation", result)
        self.assertNotIn("Footer stuff", result)

    def test_extract_text_from_pdf_invalid(self):
        # Passing invalid bytes should raise ValueError
        with self.assertRaises(ValueError):
            resume_optimizer.extract_text_from_pdf(b"invalid pdf data")

    @patch("resume_optimizer.tailor_resume")
    @patch("resume_optimizer.generate_cover_letter")
    @patch("resume_optimizer.check_ats_score")
    @patch("resume_optimizer.refine_resume")
    @patch("resume_optimizer.review_with_claude")
    @patch("resume_optimizer.humanize_with_gpt")
    def test_pipeline_looping_and_steps(self, mock_humanize, mock_claude, mock_refine, mock_ats, mock_cover_letter, mock_tailor):
        # Configure Mocks
        mock_tailor.return_value = "Draft Tailored Resume Content"
        mock_cover_letter.return_value = "Cover Letter Content"
        
        # We want to test ATS score looping:
        # First attempt: score 70 (below 85)
        # Second attempt: score 88 (above 85, terminates loop)
        mock_ats.side_effect = [
            {"score": 70, "missing_keywords": ["React"], "formatting_issues": [], "suggestions": ["Add React"]},
            {"score": 88, "missing_keywords": [], "formatting_issues": [], "suggestions": []}
        ]
        
        mock_refine.return_value = "Refined Resume Content"
        
        # Claude returns that updates were made
        mock_claude.return_value = {
            "updated": True,
            "feedback": "Claude adjusted bullet points to sound better.",
            "updated_resume": "Claude Enhanced Resume Content"
        }
        
        mock_humanize.return_value = "Final Humanized Resume Content"

        # Execute Pipeline
        results = resume_optimizer.run_optimization_pipeline(
            openai_key="mock-openai-key",
            anthropic_key="mock-claude-key",
            resume_text="Original Resume",
            jd_text="Job Description text"
        )

        # Assertions
        # 1. Tailor and Cover Letter should be called
        mock_tailor.assert_called_once()
        mock_cover_letter.assert_called_once()
        
        # 2. Check ATS score should be called twice (looping)
        self.assertEqual(mock_ats.call_count, 2)
        
        # 3. Refine should be called once (since first attempt was < 85%)
        mock_refine.assert_called_once()
        
        # 4. Claude review should be called with the refined resume
        mock_claude.assert_called_once_with("mock-claude-key", "Refined Resume Content", "Job Description text")
        
        # 5. Humanize should be called since Claude updated the resume
        mock_humanize.assert_called_once_with("mock-openai-key", "Claude Enhanced Resume Content", "Job Description text")
        
        # 6. Verify outputs structure
        self.assertEqual(results["final_resume"], "Final Humanized Resume Content")
        self.assertEqual(results["cover_letter"], "Cover Letter Content")
        self.assertEqual(results["ats_pass"], True)
        self.assertEqual(results["claude_updated"], True)
        self.assertEqual(len(results["ats_history"]), 2)
        self.assertEqual(results["ats_history"][0]["score"], 70)
        self.assertEqual(results["ats_history"][1]["score"], 88)

if __name__ == "__main__":
    unittest.main()
