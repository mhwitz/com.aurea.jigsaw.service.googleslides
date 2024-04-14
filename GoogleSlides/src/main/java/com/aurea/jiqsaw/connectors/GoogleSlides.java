package com.aurea.jiqsaw.connectors;

import com.sonicsw.esb.service.common.SFCParameters;

import com.sonicsw.esb.service.common.SFCServiceContext;
import com.sonicsw.esb.service.common.impl.AbstractSFCServiceImpl;
import com.sonicsw.xq.XQEnvelope;
import com.sonicsw.xq.XQMessage;
import com.sonicsw.xq.XQMessageException;
import com.sonicsw.xq.XQServiceException;
import org.apache.log4j.Logger;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.auth.oauth2.ServiceAccountCredentials;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;

import com.google.api.services.slides.v1.Slides;
import com.google.api.services.slides.v1.SlidesScopes;
import com.google.api.services.slides.v1.model.AffineTransform;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationRequest;
import com.google.api.services.slides.v1.model.BatchUpdatePresentationResponse;
import com.google.api.services.slides.v1.model.CreateImageRequest;
import com.google.api.services.slides.v1.model.CreateShapeRequest;
import com.google.api.services.slides.v1.model.CreateSlideRequest;
import com.google.api.services.slides.v1.model.Dimension;
import com.google.api.services.slides.v1.model.InsertTextRequest;
import com.google.api.services.slides.v1.model.LayoutReference;
import com.google.api.services.slides.v1.model.Page;
import com.google.api.services.slides.v1.model.PageElement;
import com.google.api.services.slides.v1.model.PageElementProperties;
import com.google.api.services.slides.v1.model.ParagraphStyle;
import com.google.api.services.slides.v1.model.Presentation;
import com.google.api.services.slides.v1.model.Range;
import com.google.api.services.slides.v1.model.Request;
import com.google.api.services.slides.v1.model.Response;
import com.google.api.services.slides.v1.model.Size;
import com.google.api.services.slides.v1.model.TextStyle;
import com.google.api.services.slides.v1.model.UpdateParagraphStyleRequest;
import com.google.api.services.slides.v1.model.UpdateTextStyleRequest;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


/**
 * GoogleSlides SFC Service
 */
public class GoogleSlides extends AbstractSFCServiceImpl {
	
    private  final String APPLICATION_NAME = "Google Slides";
    private  final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
    private  final String TOKENS_DIRECTORY_PATH = "tokens";
    private  final List<String> SCOPES = Arrays.asList(SlidesScopes.PRESENTATIONS, DriveScopes.DRIVE);
    private  final String CREDENTIALS_FILE_PATH = "/credentials.json";

    // access to the SFC's logging mechanism
    private final Logger log = Logger.getLogger(this.getClass());

    /**
     * Process each incoming message
     * 
     * @param _ctx runtime context of processing
     * @param _envelope contains the incoming message
     * @throws XQServiceException if the message cannot be correctly processed - message will be set to RME
     * @see com.sonicsw.esb.service.common.impl.AbstractSFCServiceImpl#doService(
     *            com.sonicsw.esb.service.common.SFCServiceContext, com.sonicsw.xq.XQEnvelope)
     */
    public final void doService(final SFCServiceContext _ctx, final XQEnvelope _envelope) throws XQServiceException {
        // get the parameters from the Service Context
        final SFCParameters parms = _ctx.getParameters();
        String securityFilePath, folderId, applicationName, presentationName, presentationId, pathToDataFile = null;
        String presenationFile = "";
		securityFilePath = parms.getParameter("SecurityFilePath", null);
		folderId = parms.getParameter("FolderId", null);
		applicationName = parms.getParameter("ApplicationName", null);
		presentationName = parms.getParameter("PresentationName",null);
		pathToDataFile = parms.getParameter("PathToFile", null);
		
        try {
			Drive driveService = getDriveService(securityFilePath, applicationName);
	        Slides slidesService = getSlideService(securityFilePath, applicationName);
	        presentationId = createPresentation(driveService, folderId, presentationName);
			createSlides(slidesService, presentationName, presentationId, applicationName, pathToDataFile);
			presenationFile = movePresentation(driveService, slidesService, presentationId, folderId);

		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (GeneralSecurityException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}


        // get the message from the envelope
        final XQMessage message = _envelope.getMessage();
        try {
			message.setStringHeader("File Link", presenationFile);
		} catch (XQMessageException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        _ctx.addIncomingToOutbox();
    }
    private static String createPresentation(Drive driveService, String folderId, String presentationName) throws IOException {
        File fileMetadata = new File();
        fileMetadata.setName(presentationName);
        fileMetadata.setMimeType("application/vnd.google-apps.presentation");
        //fileMetadata.setParents(Collections.singletonList(folderId));

        File file = driveService.files().create(fileMetadata)
                .setFields("id, webViewLink")
                .execute();

        return file.getId();
    }
    
    private static Presentation getPresentation(Slides slidesService, String presentationId) throws IOException {
        return slidesService.presentations().get(presentationId).execute();
    }
    
    private static void createSlides(Slides slidesService, String presentationName, String presentationId, String applicationName, String pathToDataFile) throws IOException, GeneralSecurityException { 	
    	
        Presentation presentation = getPresentation(slidesService, presentationId);

        // Read data from a file
        if (pathToDataFile != null) {
            List<String> slideData = readDataFromFile(pathToDataFile);

            // Create slides based on the data
            for (String data : slideData) {
                Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                        .setObjectId(presentationId)
                        .setInsertionIndex(presentation.getSlides().size())
                        .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TitleAndBody")));
                BatchUpdatePresentationResponse response  = slidesService.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                        .execute();

            }
        	
        }

        addSlideWithText(slidesService, presentationId, "Cover Page", Collections.emptyList(), true );

        //added a text box with multiple formatting

        String slide1Title = "Slide 1";
        List<String> slide0Content = Arrays.asList(
                "Main Line 1",
                "- Bullet Point 1",
                "- Bullet Point 2",
                "Main Line 2",
                "- Bullet Point 3",
                "- Bullet Point 4",
                "Main Line 3"
        );
        String slide2Title = "Slide 2";
        List<String> slide1Content = Arrays.asList(
                "Data Point 1",
                "Data Point 2",
                "Data Point 3"
        );
        String slide3Title = "Slide 3";
        List<String> slide2Content = Arrays.asList(
                "Another Data Point 1",
                "Another Data Point 2",
                "Another Data Point 3"
        );
        String slide4Title = "Slide 4";
        List<String> slide3Content = Arrays.asList(
                "More Data Point 1",
                "More Data Point 2",
                "More Data Point 3"
        );     
        
        
       
        addSlideWithText(slidesService, presentationId, slide1Title,slide0Content, false );
        addSlideWithText(slidesService, presentationId, slide2Title,slide1Content, false );
        addSlideWithText(slidesService, presentationId, slide3Title,slide2Content, false );
        addSlideWithText(slidesService, presentationId, slide4Title,slide3Content, false );
               
        //addSlideWithImage(slidesService, presentationId, "ImageURL");
        //addSlideWithImageAndText(slidesService, presentationId, "Title","slideText", "ImageURL");
        


       // addSlideWithTextBox(slidesService, presentationId, slideTitle, textBoxContent);
        
        
        // when slides are complete, move them
        
    }
    private  String movePresentation(Drive driveService, Slides slidesService, String presentationId, String folderId) throws IOException {
		File file = driveService.files().update(presentationId, null).setAddParents(folderId).execute();

	    // Get the updated presentation details
	    com.google.api.services.slides.v1.model.Presentation presentation = slidesService.presentations()
	            .get(file.getId())
	            .execute();

	    // Return the webViewLink of the moved presentation
	    return file.getWebViewLink();
	}

    private static  Slides getSlideService(String jsonPath, String applicationName) throws GeneralSecurityException, IOException {
		final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
		Slides slideService = null;
		JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();

		// Load credentials
		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

		// Load service account key from JSON file
		GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(jsonPath))
				.createScoped(Collections.singletonList("https://www.googleapis.com/auth/drive"));

		// Build a new authorized API client service
		// Build a new authorized API client service
		slideService = new Slides.Builder(httpTransport, JSON_FACTORY, credential).setApplicationName(applicationName)
				.build();
    	
    	
		return slideService;
    	
    }
	private static Drive getDriveService(String jsonPath, String applicationName)
			throws IOException, GeneralSecurityException {
		final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
		Drive driveService = null;

		JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();

		// Load credentials
		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

		// Load service account key from JSON file
		GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(jsonPath))
				.createScoped(Collections.singletonList("https://www.googleapis.com/auth/drive"));

		// Build a new authorized API client service
		driveService = new Drive.Builder(httpTransport, JSON_FACTORY, credential).setApplicationName(applicationName)
				.build();

		return driveService;
	}
  
    
    private static List<String> readDataFromFile(String filePath) throws IOException {

	    Path path = Paths.get(filePath);
	    return Files.lines(path)
	            .collect(Collectors.toList());
        // ...
    }
    private static void addSlideWithTextBox(Slides service, String presentationId, String slideTitle, List<String> textBoxContent) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add title to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = new ArrayList<>();
        requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(slideId)
                .setInsertionIndex(0)
                .setText(slideTitle)));

        // Add text box content with formatting
        StringBuilder textBoxBuilder = new StringBuilder();
        boolean isBulletPoint = false;
        for (String line : textBoxContent) {
            if (line.startsWith("- ")) {
                // Bullet point
                if (!isBulletPoint) {
                    textBoxBuilder.append("\n");
                    isBulletPoint = true;
                }
                textBoxBuilder.append(line.substring(2)).append("\n");
            } else {
                // Main line
                if (isBulletPoint) {
                    textBoxBuilder.append("\n");
                    isBulletPoint = false;
                }
                textBoxBuilder.append(line).append("\n");
            }
        }

        requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(slideId)
                .setInsertionIndex(slideTitle.length())
                .setText(textBoxBuilder.toString())));

        // Apply text formatting
        requests.add(new Request().setUpdateParagraphStyle(new UpdateParagraphStyleRequest()
                .setObjectId(slideId)
                .setTextRange(new Range().setType("ALL"))
                .setStyle(new ParagraphStyle()
                        .setIndentStart(new Dimension().setMagnitude(36.0).setUnit("PT"))
                        .setIndentFirstLine(new Dimension().setMagnitude(-36.0).setUnit("PT"))
                        .setSpaceBelow(new Dimension().setMagnitude(10.0).setUnit("PT")))));

        requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                .setObjectId(slideId)
                .setTextRange(new Range()
                        .setType("FIXED_RANGE")
                        .setStartIndex(0)
                        .setEndIndex(slideTitle.length()))
                .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))));

        int startIndex = slideTitle.length() + 1;
        for (String line : textBoxContent) {
            if (!line.startsWith("- ")) {
                // Main line
                int endIndex = startIndex + line.length();
                requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(slideId)
                        .setTextRange(new Range()
                                .setType("FIXED_RANGE")
                                .setStartIndex(startIndex)
                                .setEndIndex(endIndex))
                        .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(24.0).setUnit("PT")))));
            }
            startIndex += line.length() + 1;
        }

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }
    
    private static void addSlideWithText(Slides service, String presentationId, String slideTitle, List<String> textBoxContent, boolean isCoverPage) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout(isCoverPage ? "TITLE" : "BLANK")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Get the ID of the newly created slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();

        List<Request> requests = new ArrayList<>();

        if (isCoverPage) {
            // Add title to the cover page
            requests.add(new Request().setInsertText(new InsertTextRequest()
                    .setObjectId(slideId)
                    .setInsertionIndex(0)
                    .setText(slideTitle)));

            requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                    .setObjectId(slideId)
                    .setTextRange(new Range()
                            .setType("FIXED_RANGE")
                            .setStartIndex(0)
                            .setEndIndex(slideTitle.length()))
                    .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))
                    .setFields("*")));
        } else {
            // Create a new shape for the title
            CreateShapeRequest titleShapeRequest = new CreateShapeRequest()
                    .setShapeType("TEXT_BOX")
                    .setElementProperties(new PageElementProperties()
                            .setPageObjectId(slideId)
                            .setSize(new Size()
                                    .setWidth(new Dimension().setMagnitude(500.0).setUnit("PT"))
                                    .setHeight(new Dimension().setMagnitude(50.0).setUnit("PT")))
                            .setTransform(new AffineTransform()
                                    .setScaleX(1.0).setUnit("UNIT_PERCENT")
                                    .setScaleY(1.0).setUnit("UNIT_PERCENT")
                                    .setTranslateX(50.0).setUnit("PT")
                                    .setTranslateY(50.0).setUnit("PT")));
            requests.add(new Request().setCreateShape(titleShapeRequest));

            // Create a new text box for the content
            CreateShapeRequest textBoxShapeRequest = new CreateShapeRequest()
                    .setShapeType("TEXT_BOX")
                    .setElementProperties(new PageElementProperties()
                            .setPageObjectId(slideId)
                            .setSize(new Size()
                                    .setWidth(new Dimension().setMagnitude(500.0).setUnit("PT"))
                                    .setHeight(new Dimension().setMagnitude(200.0).setUnit("PT")))
                            .setTransform(new AffineTransform()
                                    .setScaleX(1.0).setUnit("UNIT_PERCENT")
                                    .setScaleY(1.0).setUnit("UNIT_PERCENT")
                                    .setTranslateX(50.0).setUnit("PT")
                                    .setTranslateY(100.0).setUnit("PT")));
            requests.add(new Request().setCreateShape(textBoxShapeRequest));

            // Execute the batch update to create the shapes
            BatchUpdatePresentationResponse shapeResponse = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();

            // Get the object IDs of the created shapes
            String titleId = shapeResponse.getReplies().get(0).getCreateShape().getObjectId();
            String textBoxId = shapeResponse.getReplies().get(1).getCreateShape().getObjectId();

            // Add title to the slide
            try {
                Request titleRequest = new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(titleId)
                        .setInsertionIndex(0)
                        .setText(slideTitle));
                service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(titleRequest))).execute();

                Request titleStyleRequest = new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(titleId)
                        .setTextRange(new Range()
                                .setType("FIXED_RANGE")
                                .setStartIndex(0)
                                .setEndIndex(slideTitle.length()))
                        .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))
                        .setFields("*"));
                service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(titleStyleRequest))).execute();
            } catch (GoogleJsonResponseException e) {
                // Ignore the exception if the object doesn't allow text editing
                if (!e.getDetails().getMessage().contains("does not allow text editing")) {
                    throw e;
                }
            }

            // Add text box content with formatting
            try {
                StringBuilder textBoxBuilder = new StringBuilder();
                for (int i = 0; i < textBoxContent.size(); i++) {
                    String line = textBoxContent.get(i);
                    if (i == 0) {
                        // First line with large text
                        textBoxBuilder.append(line).append("\n");
                    } else {
                        // Subsequent lines with bullets and indentation
                        textBoxBuilder.append("• ").append(line).append("\n");
                    }
                }

                Request textBoxRequest = new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(textBoxId)
                        .setInsertionIndex(0)
                        .setText(textBoxBuilder.toString()));
                service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(textBoxRequest))).execute();
            } catch (GoogleJsonResponseException e) {
                // Ignore the exception if the object doesn't allow text editing
                if (!e.getDetails().getMessage().contains("does not allow text editing")) {
                    throw e;
                }
            }
        }
    }    
    private static boolean isEditable(Slides service, String presentationId, String objectId) throws IOException {
        Presentation presentation = service.presentations().get(presentationId).execute();
        for (Page page : presentation.getSlides()) {
            for (PageElement element : page.getPageElements()) {
                if (element.getObjectId().equals(objectId)) {
                    if (element.getShape() != null && element.getShape().getPlaceholder() != null) {
                        String placeholderType = element.getShape().getPlaceholder().getType();
                        return placeholderType.equals("TITLE") || placeholderType.equals("BODY");
                    }
                    return true; // Non-placeholder shapes are considered editable
                }
            }
        }
        return false;
    }   
    
    private static void addSlideWithImage(Slides service, String presentationId, String imageUrl) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("BLANK")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add image to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = Collections.singletonList(
                new Request().setCreateImage(new CreateImageRequest()
                        .setUrl(imageUrl)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(3000000.0).setUnit("EMU"))
                                        .setHeight(new Dimension().setMagnitude(3000000.0).setUnit("EMU")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(100000.0)
                                        .setTranslateY(100000.0)
                                        .setUnit("EMU")))));

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }

    private static void addSlideWithImageAndText(Slides service, String presentationId, String slideTitle, String slideText, String imageUrl) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add title and text to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = new ArrayList<>();
        Presentation updatedPresentation = service.presentations().get(presentationId).execute();
        for (Page page : updatedPresentation.getSlides()) {
            if (page.getObjectId().equals(slideId)) {
                String titleId = null;
                String bodyId = null;
                for (PageElement element : page.getPageElements()) {
                    if (element.getShape() != null && element.getShape().getPlaceholder() != null) {
                        if (element.getShape().getPlaceholder().getType().equals("CENTERED_TITLE")) {
                            titleId = element.getObjectId();
                        } else if (element.getShape().getPlaceholder().getType().equals("BODY")) {
                            bodyId = element.getObjectId();
                        }
                    }
                }

                // Add title to the slide
                if (titleId != null) {
                    requests.add(new Request().setInsertText(new InsertTextRequest()
                            .setObjectId(titleId)
                            .setInsertionIndex(0)
                            .setText(slideTitle)));

                    requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                            .setObjectId(titleId)
                            .setTextRange(new Range()
                                    .setType("FIXED_RANGE")
                                    .setStartIndex(0)
                                    .setEndIndex(slideTitle.length()))
                            .setStyle(new TextStyle()
                                    .setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))));
                }

                // Add text to the slide
                if (bodyId != null) {
                    requests.add(new Request().setInsertText(new InsertTextRequest()
                            .setObjectId(bodyId)
                            .setInsertionIndex(0)
                            .setText(slideText)));
                }

                // Add image to the slide
                requests.add(new Request().setCreateImage(new CreateImageRequest()
                        .setUrl(imageUrl)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(3000000.0).setUnit("EMU"))
                                        .setHeight(new Dimension().setMagnitude(3000000.0).setUnit("EMU")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(0.5)
                                        .setScaleY(0.5)
                                        .setTranslateX(100000.0)
                                        .setTranslateY(3000000.0)
                                        .setUnit("EMU")))));

                break; // Exit the loop once the slide is found
            }
        }

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    } 
    
    
}