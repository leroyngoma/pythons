from fpdf import FPDF

# Create a PDF class
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'AI Readiness Report for Ciber', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

# Create a PDF object
pdf = PDF()
pdf.add_page()

# Set font for content
pdf.set_font('Arial', '', 12)

# Add content to the PDF
report_content = """
**Prepared for:** Leroy Ngoma
**Organization Size:** 11-50

---

### Executive Summary

As organizations across industries increasingly embrace artificial intelligence (AI), understanding their readiness to implement AI solutions is essential. This comprehensive report evaluates Ciber's current position regarding AI readiness based on survey responses. The findings indicate that while Ciber has identified a specific area for improvement through AI, significant gaps exist in understanding, data management, expertise, project management, investment strategy, and team confidence.

To leverage AI effectively, Ciber must prioritize addressing these gaps. AiBizHive is uniquely positioned to assist Ciber in this transformative journey, providing tailored solutions and expertise to build a strong foundation for AI adoption.

---

### Key Takeaways for AiBizHive

1. **Education on AI Applications:**
   - Ciber lacks awareness of how AI can address business challenges. Conducting educational workshops on AI fundamentals and applications can significantly enhance their understanding and 
openness to AI solutions.

2. **Data Management Improvement:**
   - Ciber currently does not collect or store data effectively. Offering services to develop a robust data strategy and management practices will be crucial in preparing them for AI initiatives.

3. **Building AI Expertise:**
   - With no in-house AI expertise or partnerships, Ciber requires support in this area. AiBizHive can connect them with AI consultants and training programs to build necessary capabilities.

4. **Structured Project Management:**
   - The absence of a structured approach to AI project management indicates a significant readiness gap. Providing a project management framework tailored for AI initiatives can facilitate successful implementation.

5. **Investment Strategy Development:**
   - Ciber has not allocated any budget for AI investments. AiBizHive can assist in conducting cost-benefit analyses and identifying funding opportunities to support initial projects.      

6. **Confidence Building for AI Execution:**
   - The team’s low confidence in executing AI projects presents a barrier to successful adoption. Offering skill development workshops and mentorship opportunities can empower Ciber's team.

---

### Detailed Analysis

#### 1. Understanding of AI and Its Applications
- **Score:** 1 (Lowest Confidence)
- **Observation:** Ciber is unaware of how AI could address its specific business challenges, limiting its ability to consider AI as a viable solution.
- **Recommendation:** AiBizHive can conduct AI awareness workshops tailored to Ciber's needs, utilizing industry-specific case studies to illustrate the practical applications of AI.       

#### 2. Data Readiness
- **Score:** 1 (Insufficient Data Practices)
- **Observation:** Ciber lacks effective data collection and storage practices, making it challenging to leverage AI insights.
- **Recommendation:** AiBizHive should assist Ciber in developing a data management framework that includes tools for efficient data collection, storage, and organization.

#### 3. AI Expertise and Resources
- **Score:** 1 (No Existing Expertise)
- **Observation:** Ciber currently lacks in-house AI expertise and partnerships, which may hinder progress.
- **Recommendation:** AiBizHive can facilitate connections with AI consultants and training programs to build a foundational understanding of AI within Ciber.

#### 4. AI Project Management and Implementation
- **Score:** 1 (No Project Management Plans)
- **Observation:** Ciber lacks structured project management plans for AI initiatives, indicating unpreparedness for implementation.
- **Recommendation:** AiBizHive can provide a structured project management framework specifically for AI projects, encompassing the entire lifecycle from ideation to execution.
- **Score:** 1 (No Project Management Plans)
- **Observation:** Ciber lacks structured project management plans for AI initiatives, indicating unpreparedness for implementation.
- **Recommendation:** AiBizHive can provide a structured project management framework specifically for AI projects, encompassing the entire lifecycle from ideation to execution.

#### 5. AI Investment Approach

#### 5. AI Investment Approach
- **Score:** 1 (No Budget Allocation)
- **Observation:** Ciber has not allocated any budget for AI investments, demonstrating caution.
- **Recommendation:** AiBizHive should assist Ciber in conducting cost-benefit analyses of potential AI projects and identifying funding opportunities to support initial investments.       

#### 6. Team Confidence in AI Execution
- **Score:** 1 (Low Confidence)
- **Observation:** Ciber's team exhibits low confidence in executing AI projects, posing a barrier to adoption.
- **Recommendation:** AiBizHive can offer workshops aimed at enhancing team skills in AI implementation and project management, as well as mentorship programs with experienced AI practitioners.

---

### Conclusion

Ciber is at a critical juncture regarding AI readiness, with several gaps that must be addressed to harness the full potential of AI. By leveraging AiBizHive's expertise, Ciber can build a 
solid foundation for AI initiatives that align with its strategic goals.

### Recommended Next Steps

1. **Schedule AI Awareness Workshops:** Initiate discussions to set a date for workshops that enhance understanding of AI applications.
2. **Develop a Comprehensive Data Strategy:** Assist Ciber in mapping out a data management framework that supports effective data collection and storage.
3. **Explore Training Opportunities:** Identify and enroll team members in relevant training sessions to build foundational AI skills.
4. **Plan Initial Pilot Projects:** Brainstorm suitable pilot projects that can serve as a first step toward AI implementation.
5. **Conduct Cost-Benefit Analyses:** Aid Ciber in assessing the financial implications of potential AI projects to justify investment.

By focusing on these actionable steps, Ciber can position itself for successful AI integration, ultimately leading to improved operational efficiency and competitive advantage in the marketplace.
"""

# Replace problematic characters with simple quotes and dashes
report_content = report_content.replace("’", "'").replace("“", '"').replace("”", '"')

# Split content into lines and add to PDF
for line in report_content.split('\n'):
    pdf.multi_cell(0, 10, line)

# Save the PDF to a file
pdf.output('AI_Readiness_Report_Ciber2.pdf')
