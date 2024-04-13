package com.aurea.jiqsaw.connectors;

import com.sonicsw.esb.service.common.SFCParameters;

import com.sonicsw.esb.service.common.SFCServiceContext;
import com.sonicsw.esb.service.common.impl.AbstractSFCServiceImpl;
import com.sonicsw.xq.XQEnvelope;
import com.sonicsw.xq.XQMessage;
import com.sonicsw.xq.XQServiceException;
import org.apache.log4j.Logger;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
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
import com.google.api.services.slides.v1.model.CreateSlideRequest;
import com.google.api.services.slides.v1.model.Dimension;
import com.google.api.services.slides.v1.model.InsertTextRequest;
import com.google.api.services.slides.v1.model.LayoutReference;
import com.google.api.services.slides.v1.model.Page;
import com.google.api.services.slides.v1.model.PageElementProperties;
import com.google.api.services.slides.v1.model.ParagraphStyle;
import com.google.api.services.slides.v1.model.Presentation;
import com.google.api.services.slides.v1.model.Range;
import com.google.api.services.slides.v1.model.Request;
import com.google.api.services.slides.v1.model.Size;
import com.google.api.services.slides.v1.model.TextStyle;
import com.google.api.services.slides.v1.model.UpdateParagraphStyleRequest;
import com.google.api.services.slides.v1.model.UpdateTextStyleRequest;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.Map;


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
        String securityFilePath, baseFolderId, applicationName = null;
		securityFilePath = parms.getParameter("securityFilePath");
		baseFolderId = parms.getParameter("baseFolderId");
		applicationName = parms.getParameter("applicationName");
		String slideName = parms.getParameter("SlideName");
		try {
			createSlides(securityFilePath, slideName, baseFolderId, applicationName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (GeneralSecurityException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // get the message from the envelope
        final XQMessage message = _envelope.getMessage();
        
        _ctx.addIncomingToOutbox();
    }
   
    private void createSlides(String securityFilePath, String slideName, String baseFolderId, String applicationName) throws IOException, GeneralSecurityException { 	
    	
        Drive driveService = getDriveService(securityFilePath, applicationName);


        Slides slidesService = getSlideService(securityFilePath, applicationName);

        String presentationId = "YOUR_PRESENTATION_ID";
        Presentation presentation = slidesService.presentations().get(presentationId).execute();

        // Read data from a file
        List<String> slideData = readDataFromFile("path/to/your/file.txt");

        // Create slides based on the data
        for (String data : slideData) {
            Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                    .setObjectId(presentationId)
                    .setInsertionIndex(presentation.getSlides().size())
                    .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));
            BatchUpdatePresentationResponse response  = slidesService.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                    .execute();

        }
        
        addSlideWithText(slidesService, presentationId, "Title","slideText" );
        addSlideWithImage(slidesService, presentationId, "ImageURL");
        addSlideWithImageAndText(slidesService, presentationId, "Title","slideText", "ImageURL");
        
        //added a text box with multiple formatting
        String slideTitle = "My Slide Title";
        List<String> textBoxContent = Arrays.asList(
                "Main Line 1",
                "- Bullet Point 1",
                "- Bullet Point 2",
                "Main Line 2",
                "- Bullet Point 3",
                "- Bullet Point 4",
                "Main Line 3"
        );

        addSlideWithTextBox(slidesService, presentationId, slideTitle, textBoxContent);
        
        
        // when slides are complete, move them
        
		try {
			driveService.files().update(presentationId, null).setAddParents(baseFolderId).execute();
		} catch (Exception e) {
			e.printStackTrace();
		}

    }
    public Slides getSlideService(String jsonPath, String applicationName) throws GeneralSecurityException, IOException {
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
	public static Drive getDriveService(String jsonPath, String applicationName)
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
  
    
    private  List<String> readDataFromFile(String filePath) throws IOException {
		return SCOPES;
        // Implement the logic to read data from the file and return it as a list of strings
        // ...
    }
    private static void addSlideWithTextBox(Slides service, String presentationId, String slideTitle, List<String> textBoxContent) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(presentationId)
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
                .setTextRange(new Range().setType("SPECIFIC_RANGE").setStartIndex(0).setEndIndex(slideTitle.length()))
                .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))));

        int startIndex = slideTitle.length() + 1;
        for (String line : textBoxContent) {
            if (!line.startsWith("- ")) {
                // Main line
                int endIndex = startIndex + line.length();
                requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(slideId)
                        .setTextRange(new Range().setType("SPECIFIC_RANGE").setStartIndex(startIndex).setEndIndex(endIndex))
                        .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(24.0).setUnit("PT")))));
            }
            startIndex += line.length() + 1;
        }

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }    
    
    private static void addSlideWithText(Slides service, String presentationId, String slideTitle, String slideText) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(presentationId)
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));
        
        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add title and text to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = Arrays.asList(
                new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(slideId)
                        .setInsertionIndex(0)
                        .setText(slideTitle)),
                new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(slideId)
                        .setInsertionIndex(slideTitle.length())
                        .setText("\n" + slideText)));

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }

    private static void addSlideWithImage(Slides service, String presentationId, String imageUrl) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(presentationId)
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add image to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = Collections.singletonList(
                new Request().setCreateImage(new CreateImageRequest()
                        .setUrl(imageUrl)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size().setWidth(new Dimension().setMagnitude(3000000.0).setUnit("EMU"))
                                        .setHeight(new Dimension().setMagnitude(3000000.0).setUnit("EMU")))
                                .setTransform(new AffineTransform().setScaleX(1.0).setScaleY(1.0).setTranslateX(100000.0).setTranslateY(100000.0).setUnit("EMU")))));

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }

    private static void addSlideWithImageAndText(Slides service, String presentationId, String slideTitle, String slideText, String imageUrl) throws IOException {
        // Create a new slide
        Presentation presentation = service.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(presentationId)
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE_AND_BODY")));

        BatchUpdatePresentationResponse response = service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Add title and text to the slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();
        List<Request> requests = Arrays.asList(
                new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(slideId)
                        .setInsertionIndex(0)
                        .setText(slideTitle)),
                new Request().setInsertText(new InsertTextRequest()
                        .setObjectId(slideId)
                        .setInsertionIndex(slideTitle.length())
                        .setText("\n" + slideText)),
                new Request().setCreateImage(new CreateImageRequest()
                        .setUrl(imageUrl)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size().setWidth(new Dimension().setMagnitude(3000000.0).setUnit("EMU"))
                                        .setHeight(new Dimension().setMagnitude(3000000.0).setUnit("EMU")))
                                .setTransform(new AffineTransform().setScaleX(0.5).setScaleY(0.5).setTranslateX(100000.0).setTranslateY(3000000.0).setUnit("EMU")))));

        service.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    } 	
}