package com.aurea.jiqsaw.connectors;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.CopySheetToAnotherSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.model.AffineTransform;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationRequest;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationResponse;
import com.google.api.services.slides.v1.model.CreateShapeRequest;
import com.google.api.services.slides.v1.model.CreateSlideRequest;
import com.google.api.services.slides.v1.model.Dimension;
import com.google.api.services.slides.v1.model.InsertTextRequest;
import com.google.api.services.slides.v1.model.LayoutReference;
import com.google.api.services.slides.v1.model.Page;
import com.google.api.services.slides.v1.model.PageElement;
import com.google.api.services.slides.v1.model.PageElementProperties;
import com.google.api.services.slides.v1.model.Presentation;
import com.google.api.services.slides.v1.model.ReplaceAllTextRequest;
import com.google.api.services.slides.v1.model.Request;
import com.google.api.services.slides.v1.model.Shape;
import com.google.api.services.slides.v1.model.Size;
import com.google.api.services.slides.v1.model.SubstringMatchCriteria;

public class ProductHub {
    private static String TEMPLATE_PRESENTATION_ID = "your-template-presentation-id";
    private static  String TITLE_TEXT_BOX_ID = "title-text-box-object-id";
    private static  String DATE_TEXT_BOX_ID = "date-text-box-object-id";
    private static  String SLIDE2_TITLE_TEXT_BOX_ID = "slide2-title-text-box-object-id";
    private static  String SLIDE2_TEXT_BOX_ID = "slide2-text-box-object-id";
	
	public ProductHub (String TPI, String TTBI, String DTTI, String STTBI, String STBI) {
		TEMPLATE_PRESENTATION_ID = TPI;
		TITLE_TEXT_BOX_ID = TTBI;
		DATE_TEXT_BOX_ID = DTTI;
		SLIDE2_TITLE_TEXT_BOX_ID = STTBI;
		SLIDE2_TEXT_BOX_ID = STBI;
	}
    public String getGeneratedContent(Sheets sheetsService, String spreadsheetId, String pageName, String cellAddress) throws IOException {
        // Construct the range in A1 notation
        String range = pageName + "!" + cellAddress;

        // Retrieve the formatted value from the specified cell
        ValueRange response = sheetsService.spreadsheets().values()
                .get(spreadsheetId, range)
                .setValueRenderOption("FORMATTED_VALUE")
                .execute();

        // Get the generated content from the response
        List<List<Object>> values = response.getValues();
        if (values == null || values.isEmpty()) {
            System.out.println("No data found in the specified cell.");
            return null;
        } else {
            String generatedContent = values.get(0).get(0).toString();
            //System.out.println("Generated content: " + generatedContent);
            return generatedContent;
        }
    }
   
	
	public String getValueFromSheet(Sheets sheetsService, String spreadsheetId, String pageName, String cellAddress) throws IOException {
	        // Construct the range in A1 notation
	        String range = pageName + "!" + cellAddress;

	        // Retrieve the value from the specified cell
	        ValueRange response = sheetsService.spreadsheets().values()
	                .get(spreadsheetId, range)
	                .execute();

	        // Get the value from the response
	        List<List<Object>> values = response.getValues();
	        if (values == null || values.isEmpty()) {
	            System.out.println("No data found in the specified cell.");
	            return null;
	        } else {
	            String value = values.get(0).get(0).toString();
	            System.out.println("Value retrieved: " + value);
	            return value;
	        }
	    }
	   
	    public String createPresentationFromTemplate(Slides slidesService, String title, List<Slide2Data> slide2DataList) throws IOException {
	        // Create a new presentation
	        Presentation newPresentation = new Presentation()
	                .setTitle(title);
	        newPresentation = slidesService.presentations().create(newPresentation).execute();
	        String newPresentationId = newPresentation.getPresentationId();

	        // Retrieve the template presentation
	        Presentation templatePresentation = slidesService.presentations().get(TEMPLATE_PRESENTATION_ID).execute();

	        // Create requests to copy slides from the template presentation
	        List<Request> requests = new ArrayList<>();

	        // Copy Slide 1 from the template presentation
	        String slide1ObjectId = templatePresentation.getSlides().get(0).getObjectId();
	        requests.add(new Request()
	                .setCreateSlide(new CreateSlideRequest()
	                        .setObjectId(slide1ObjectId)
	                        .setInsertionIndex(0)));

	        // Update the Title and Date text boxes on Slide 1
	        requests.add(new Request()
	                .setReplaceAllText(new ReplaceAllTextRequest()
	                        .setContainsText(new SubstringMatchCriteria()
	                                .setText("{{Title}}")
	                                .setMatchCase(true))
	                        .setReplaceText(title)
	                        .setPageObjectIds(Collections.singletonList(slide1ObjectId))));

	        requests.add(new Request()
	                .setReplaceAllText(new ReplaceAllTextRequest()
	                        .setContainsText(new SubstringMatchCriteria()
	                                .setText("{{Date}}")
	                                .setMatchCase(true))
	                        .setReplaceText(LocalDate.now().toString())
	                        .setPageObjectIds(Collections.singletonList(slide1ObjectId))));

	        // Copy Slide 2 from the template presentation and duplicate it for each data entry
	        String slide2ObjectId = templatePresentation.getSlides().get(1).getObjectId();
	        for (int i = 0; i < slide2DataList.size(); i++) {
	            Slide2Data slide2Data = slide2DataList.get(i);
	            String duplicatedSlideObjectId = String.format("slide2_%d", i);

	            requests.add(new Request()
	                    .setCreateSlide(new CreateSlideRequest()
	                            .setObjectId(duplicatedSlideObjectId)
	                            .setInsertionIndex(i + 1)));

	            // Update the Title and Text text boxes on the duplicated Slide 2
	            requests.add(new Request()
	                    .setReplaceAllText(new ReplaceAllTextRequest()
	                            .setContainsText(new SubstringMatchCriteria()
	                                    .setText("{{Slide2Title}}")
	                                    .setMatchCase(true))
	                            .setReplaceText(slide2Data.getTitle())
	                            .setPageObjectIds(Collections.singletonList(duplicatedSlideObjectId))));

	            requests.add(new Request()
	                    .setReplaceAllText(new ReplaceAllTextRequest()
	                            .setContainsText(new SubstringMatchCriteria()
	                                    .setText("{{Slide2Text}}")
	                                    .setMatchCase(true))
	                            .setReplaceText(slide2Data.getText())
	                            .setPageObjectIds(Collections.singletonList(duplicatedSlideObjectId))));
	        }

	        // Execute the batch update to create the presentation
	        slidesService.presentations().batchUpdate(newPresentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

	        return newPresentationId;
	        
	        
	    }
	    public static String createPresentationFromTemplate2(Slides slidesService, String title, SlideData[] slideDataArray) throws IOException {
	        // Retrieve the template presentation
	        Presentation templatePresentation = slidesService.presentations().get(TEMPLATE_PRESENTATION_ID).execute();

	        // Create a new presentation
	        Presentation newPresentation = new Presentation()
	                .setTitle(title);
	        newPresentation = slidesService.presentations().create(newPresentation).execute();
	        String newPresentationId = newPresentation.getPresentationId();

	        // Create requests to modify and create slides
	        List<Request> requests = new ArrayList<>();

	        // Add the first page (Slide 1) from the template
	        String slide1ObjectId = templatePresentation.getSlides().get(1).getObjectId();
	        requests.add(new Request()
	                .setCreateSlide(new CreateSlideRequest()
	                        .setObjectId(slide1ObjectId)
	                        .setInsertionIndex(0)));

	        // Update the Title and Date fields on Slide 1
	        requests.add(new Request()
	                .setReplaceAllText(new ReplaceAllTextRequest()
	                        .setContainsText(new SubstringMatchCriteria()
	                                .setText("{{Title}}"))
	                        .setReplaceText(title)
	                        .setPageObjectIds(Collections.singletonList(slide1ObjectId))));

	        String currentDate = LocalDate.now().format(DateTimeFormatter.ofPattern("MMMM d, yyyy"));
	        requests.add(new Request()
	                .setReplaceAllText(new ReplaceAllTextRequest()
	                        .setContainsText(new SubstringMatchCriteria()
	                                .setText("{{Date}}"))
	                        .setReplaceText(currentDate)
	                        .setPageObjectIds(Collections.singletonList(slide1ObjectId))));

	        // Create subsequent pages (Slide 3) for each SlideData entry
	        if (!(slideDataArray.length > 0)) {
		        slidesService.presentations().batchUpdate(newPresentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

		        return newPresentationId;
	        	
	        }
	        for (int i = 0; i < slideDataArray.length; i++) {
	            SlideData slideData = slideDataArray[i];
	            String slideObjectId = "Slide3_" + System.currentTimeMillis();

	            requests.add(new Request()
	                    .setCreateSlide(new CreateSlideRequest()
	                            .setObjectId(slideObjectId)
	                            .setInsertionIndex(i + 1)
	                            .setSlideLayoutReference(new LayoutReference()
	                                    .setLayoutId(templatePresentation.getSlides().get(3).getObjectId()))));

	            // Update the Title and Data fields on the new slide
	            requests.add(new Request()
	                    .setReplaceAllText(new ReplaceAllTextRequest()
	                            .setContainsText(new SubstringMatchCriteria()
	                                    .setText("{{Slide Title}}"))
	                            .setReplaceText(slideData.getPageTitleValue())
	                            .setPageObjectIds(Collections.singletonList(slideObjectId))));

	            requests.add(new Request()
	                    .setReplaceAllText(new ReplaceAllTextRequest()
	                            .setContainsText(new SubstringMatchCriteria()
	                                    .setText("{{Slide Data}}"))
	                            .setReplaceText(slideData.getTextBoxValue())
	                            .setPageObjectIds(Collections.singletonList(slideObjectId))));
	        }

	        // Execute the batch update to create the presentation
	        slidesService.presentations().batchUpdate(newPresentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

	        return newPresentationId;
	    }	
	    
	    public static class Slide2Data {
	        private String title;
	        private String text;

	        public Slide2Data(String title, String text) {
	            this.title = title;
	            this.text = text;
	        }

	        public String getTitle() {
	            return title;
	        }

	        public String getText() {
	            return text;
	        }
	    }
    public void createSlideFromSheet(Slides slidesService, Sheets sheetsService, String spreadsheetId, String range, String presentationId, String slideTitle)
            throws IOException, GeneralSecurityException {

        // Retrieve data from Sheets
        List<List<Object>> values = sheetsService.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute()
                .getValues();

        if (values == null || values.isEmpty()) {
            System.out.println("No data found in the specified range.");
            return;
        }

        List<String> textBoxContent = new ArrayList<>();
        for (List<Object> row : values) {
            for (Object cell : row) {
                textBoxContent.add(String.valueOf(cell));
            }
        }

        // Create a new slide
        List<Request> requests = new ArrayList<>();
        String slideId = "slide_" + System.currentTimeMillis();

        // Create Slide
        requests.add(new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(slideId)
                .setInsertionIndex(1) // Adjust based on where you want the slide
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("BLANK"))));

        // Create title shape
        String titleId = "title_" + System.currentTimeMillis();
        requests.add(new Request().setCreateShape(new CreateShapeRequest()
                .setObjectId(titleId)
                .setShapeType("TEXT_BOX")
                .setElementProperties(new PageElementProperties()
                        .setPageObjectId(slideId)
                        .setSize(new Size()
                                .setWidth(new Dimension().setMagnitude(600.0).setUnit("PT"))
                                .setHeight(new Dimension().setMagnitude(50.0).setUnit("PT")))
                        .setTransform(new AffineTransform()
                                .setScaleX(1.0)
                                .setScaleY(1.0)
                                .setTranslateX(100.0)
                                .setTranslateY(50.0)
                                .setUnit("PT")))));

        // Insert title text
        requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(titleId)
                .setInsertionIndex(0)
                .setText(slideTitle)));

        // Create textbox for content
        String textBoxId = "textBox_" + System.currentTimeMillis();
        requests.add(new Request().setCreateShape(new CreateShapeRequest()
                .setObjectId(textBoxId)
                .setShapeType("TEXT_BOX")
                .setElementProperties(new PageElementProperties()
                        .setPageObjectId(slideId)
                        .setSize(new Size()
                                .setWidth(new Dimension().setMagnitude(600.0).setUnit("PT"))
                                .setHeight(new Dimension().setMagnitude(300.0).setUnit("PT")))
                        .setTransform(new AffineTransform()
                                .setScaleX(1.0)
                                .setScaleY(1.0)
                                .setTranslateX(100.0)
                                .setTranslateY(110.0)
                                .setUnit("PT")))));

        // Insert text content
        StringBuilder textBoxBuilder = new StringBuilder();
        for (String line : textBoxContent) {
            textBoxBuilder.append(line).append("\n");
        }
        requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(textBoxId)
                .setInsertionIndex(0)
                .setText(textBoxBuilder.toString())));

        // Execute all requests in one batch update
        BatchUpdatePresentationResponse response = slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        System.out.println("Slide created with ID: " + slideId);
    }
    
    public class SlideData {
        private String pageTitleValue;
        private String textBoxValue;

        public SlideData(String pageTitleValue, String textBoxValue) {
            this.pageTitleValue = pageTitleValue;
            this.textBoxValue = textBoxValue;
        }

        public String getPageTitleValue() {
            return pageTitleValue;
        }

        public String getTextBoxValue() {
            return textBoxValue;
        }
    }
    public SlideData[] parseSlides2(String sheetValues) {
        List<SlideData> slideDataList = new ArrayList<>();
        String[] slides = sheetValues.split("\\r\\n\\r\\n");

        for (String slide : slides) {
            String[] lines = slide.split("\\r\\n");
            String title = lines[0].split(": ")[1];
            StringBuilder textBoxBuilder = new StringBuilder();

            for (int i = 1; i < lines.length; i++) {
                String line = lines[i];
                if (line.startsWith("- ")) {
                    textBoxBuilder.append("• ").append(line.substring(2)).append("\n");
                }
            }
            SlideData slideData = new SlideData(title, textBoxBuilder.toString().trim());
            slideDataList.add(slideData);
        }

        return slideDataList.toArray(new SlideData[0]);
    }

    public SlideData[] parseSlides(String dataString) {
            List<SlideData> slideDataList = new ArrayList<>();

            // Split the data string into slide sections
            String[] slideSections = dataString.split("(?=Slide \\d+:)");

            // Iterate over each slide section
            for (String slideSection : slideSections) {
                // Extract the slide title and text box values using regular expressions
                Pattern titlePattern = Pattern.compile("Slide \\d+: (.+)");
                Pattern textBoxPattern = Pattern.compile("(?:- (.+)\\n?)+");

                Matcher titleMatcher = titlePattern.matcher(slideSection);
                Matcher textBoxMatcher = textBoxPattern.matcher(slideSection);

                if (titleMatcher.find() && textBoxMatcher.find()) {
                    String pageTitleValue = titleMatcher.group(1).trim();
                    String textBoxValue = textBoxMatcher.group(0).trim().replace("- ", "• ");

                    // Create a SlideData object and add it to the list
                    SlideData slideData = new SlideData(pageTitleValue, textBoxValue);
                    slideDataList.add(slideData);
                }
            }

            // Convert the list to an array and return it
            return slideDataList.toArray(new SlideData[0]);
        }
//    public static String createInitialNewPresentationFromTemplate(Slides slidesService, String title) throws IOException {
//        Presentation templatePresentation = slidesService.presentations().get(TEMPLATE_PRESENTATION_ID).execute();
//
//        Presentation newPresentation = new Presentation()
//                .setTitle(title);
//        newPresentation = slidesService.presentations().create(newPresentation).execute();
//
//        // Copy the slides from the template presentation to the new presentation
//        List<Request> requests = new ArrayList<>();
//        for (Page page : templatePresentation.getSlides()) {
//        	if (page.getObjectId().length()  < 6) continue;
//            requests.add(new Request().setCreateSlide(new CreateSlideRequest()
//                    .setObjectId(page.getObjectId())
//                    .setInsertionIndex(requests.size())));
//        }
//
//        // Execute the batch update to copy the slides
//        BatchUpdatePresentationResponse response = slidesService.presentations().batchUpdate(newPresentation.getPresentationId(), new BatchUpdatePresentationRequest().setRequests(requests)).execute();
//
//        String newPresentationId = newPresentation.getPresentationId();
//
//        // Wait for the changes to propagate
//        int retries = 0;
//        int maxRetries = 5;
//        int delay = 2000; // 2 seconds
//        while (retries < maxRetries) {
//            try {
//                newPresentation = slidesService.presentations().get(newPresentationId).execute();
//                if (newPresentation.getSlides().size() > 0) {
//                    break;
//                }
//            } catch (Exception e) {
//                // Ignore the exception and retry
//            }
//            retries++;
//            try {
//                Thread.sleep(delay);
//            } catch (InterruptedException e) {
//                e.printStackTrace();
//            }
//        }
//
//        // Create requests to modify slides
//        requests.clear();
//    	
//    	return newPresentationId;
//    }
    public Presentation createInitialPresentaton(Slides slidesService, String title) throws IOException {
        Presentation templatePresentation = slidesService.presentations().get(TEMPLATE_PRESENTATION_ID).execute();

        Presentation newPresentation = new Presentation()
                .setTitle(title);
        newPresentation = slidesService.presentations().create(newPresentation).execute();

        // Copy the slides from the template presentation to the new presentation
        List<Request> requests = new ArrayList<>();
        for (Page page : templatePresentation.getSlides()) {
        	if (page.getObjectId().length()  < 6) continue;
            requests.add(new Request().setCreateSlide(new CreateSlideRequest()
                    .setObjectId(page.getObjectId())
                    .setInsertionIndex(requests.size())));
        }

        // Execute the batch update to copy the slides
        BatchUpdatePresentationResponse response = slidesService.presentations().batchUpdate(newPresentation.getPresentationId(), new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        String newPresentationId = newPresentation.getPresentationId();

        // Wait for the changes to propagate
        int retries = 0;
        int maxRetries = 5;
        int delay = 2000; // 2 seconds
        while (retries < maxRetries) {
            try {
                newPresentation = slidesService.presentations().get(newPresentationId).execute();
                if (newPresentation.getSlides().size() > 0) {
                    break;
                }
            } catch (Exception e) {
                // Ignore the exception and retry
            }
            retries++;
            try {
                Thread.sleep(delay);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        // Create requests to modify slides
        requests.clear();

    	return newPresentation;
    }
    public File copyAndRenameSpreadsheet(Drive driveService, Sheets sheetsService, String spreadsheetId, String folderId, String productName) throws IOException {
        // Copy the spreadsheet to the specified folder
   
        File fileMetadata = new File();
        fileMetadata.setParents(Collections.singletonList(folderId));
        File copiedFile = driveService.files().copy(spreadsheetId, fileMetadata)
                .setFields("id, name, webViewLink")
                .execute();

        // Rename the copied spreadsheet
        File renamedFile = new File();
        renamedFile.setName(productName + " AI Roadmap");
        File updatedFile = driveService.files().update(copiedFile.getId(), renamedFile)
                .setFields("id, name, webViewLink")
                .execute();

//        // Copy the sheets from the original spreadsheet to the new spreadsheet
//        CopySheetToAnotherSpreadsheetRequest copyRequest = new CopySheetToAnotherSpreadsheetRequest()
//                .setDestinationSpreadsheetId(updatedFile.getId());
//        sheetsService.spreadsheets().sheets()
//                .copyTo(spreadsheetId, null, copyRequest)
//                .execute();

        return updatedFile;
    }

    public String updatePresentation(String newPresentationId, Slides slidesService, String title, SlideData[] slideDataArray) throws IOException {
        // Retrieve the template presentation
        List<Request> requests = new ArrayList<>();
        Presentation newPresentation = slidesService.presentations().get(newPresentationId).execute();

        String titleSlideObjectId = newPresentation.getSlides().get(0).getObjectId();
        //examineSlideTextElements(slidesService, newPresentationId, titleSlideObjectId);
        
        requests.add(new Request()
                .setReplaceAllText(new ReplaceAllTextRequest()
                        .setContainsText(new SubstringMatchCriteria()
                                .setText("{{Title}}"))
                        .setReplaceText(title)
                        .setPageObjectIds(Collections.singletonList(titleSlideObjectId))));

        String currentDate = LocalDate.now().format(DateTimeFormatter.ofPattern("MMMM d, yyyy"));
        requests.add(new Request()
                .setReplaceAllText(new ReplaceAllTextRequest()
                        .setContainsText(new SubstringMatchCriteria()
                                .setText("{{Date}}"))
                        .setReplaceText(currentDate)
                        .setPageObjectIds(Collections.singletonList(titleSlideObjectId))));

        slidesService.presentations().batchUpdate(newPresentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        requests.clear();

               
         for (int i = 0; i < slideDataArray.length; i++) {
            SlideData slideData = slideDataArray[i];
            String SlideObjectId = newPresentation.getSlides().get(i+1).getObjectId();

            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{Slide Title}}"))
                            .setReplaceText(slideData.getPageTitleValue())
                            .setPageObjectIds(Collections.singletonList(SlideObjectId))));

            requests.add(new Request()
                    .setReplaceAllText(new ReplaceAllTextRequest()
                            .setContainsText(new SubstringMatchCriteria()
                                    .setText("{{Slide Data}}"))
                            .setReplaceText(slideData.getTextBoxValue())
                            .setPageObjectIds(Collections.singletonList(SlideObjectId))));
        }

        // Execute the batch update to modify the presentation
        slidesService.presentations().batchUpdate(newPresentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        return newPresentationId;
    }
    
    
}