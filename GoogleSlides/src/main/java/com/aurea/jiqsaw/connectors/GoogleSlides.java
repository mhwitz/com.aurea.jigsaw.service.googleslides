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
//import com.google.api.client.json.gson.GsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.auth.http.HttpCredentialsAdapter;
import com.google.auth.oauth2.GoogleCredentials;
import com.google.auth.oauth2.ServiceAccountCredentials;

import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.Permission;
import com.google.api.services.drive.model.PermissionList;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
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
import com.google.api.services.slides.v1.model.CreateSlideResponse;
import com.google.api.services.slides.v1.model.Dimension;
import com.google.api.services.slides.v1.model.DuplicateObjectRequest;
import com.google.api.services.slides.v1.model.InsertTextRequest;
import com.google.api.services.slides.v1.model.LayoutReference;
import com.google.api.services.slides.v1.model.OpaqueColor;
import com.google.api.services.slides.v1.model.OptionalColor;
import com.google.api.services.slides.v1.model.Page;
import com.google.api.services.slides.v1.model.PageElement;
import com.google.api.services.slides.v1.model.PageElementProperties;
import com.google.api.services.slides.v1.model.ParagraphMarker;
import com.google.api.services.slides.v1.model.ParagraphStyle;
import com.google.api.services.slides.v1.model.Presentation;
import com.google.api.services.slides.v1.model.Range;
import com.google.api.services.slides.v1.model.Request;
import com.google.api.services.slides.v1.model.Response;
import com.google.api.services.slides.v1.model.RgbColor;
import com.google.api.services.slides.v1.model.Shape;
import com.google.api.services.slides.v1.model.ShapeBackgroundFill;
import com.google.api.services.slides.v1.model.ShapeProperties;
import com.google.api.services.slides.v1.model.Size;
import com.google.api.services.slides.v1.model.SolidFill;
import com.google.api.services.slides.v1.model.TextContent;
import com.google.api.services.slides.v1.model.TextElement;
import com.google.api.services.slides.v1.model.TextRun;
import com.google.api.services.slides.v1.model.TextStyle;
import com.google.api.services.slides.v1.model.UpdateParagraphStyleRequest;
import com.google.api.services.slides.v1.model.UpdateShapePropertiesRequest;
import com.google.api.services.slides.v1.model.UpdateSlidesPositionRequest;
import com.google.api.services.slides.v1.model.UpdateTextStyleRequest;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.GeneralSecurityException;
import java.time.temporal.ValueRange;
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
	
	private static String imageID = "https://drive.google.com/uc?id=13KZOOhxpEb7H0MYxQGjS9EMXk7k2owZE";
    private static final String LOGO_IMAGE_URL = imageID;
	private static final String RIGHT_IMAGE_URL =imageID;
	private static final String TOP_RIGHT_IMAGE_URL = imageID;
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
        String securityFilePath, folderId, applicationName, presentationName, presentationId, dataLocation, sheetRange = null;
        int dataType = 0;
        String presenationFile = "";
        Drive driveService = null;
        Sheets sheetsService = null;
        Slides slidesService = null;
        String templateName = null;
        
		securityFilePath = parms.getParameter("SecurityFilePath", null);
		folderId = parms.getParameter("FolderId", null);
		applicationName = parms.getParameter("ApplicationName", null);
		presentationName = parms.getParameter("PresentationName",null);
		dataType = parms.getIntParameter("DataType");
		dataLocation = parms.getParameter("DataLocation", null);
		sheetRange = parms.getParameter("SheetRange", null);
		String deleteFile = parms.getParameter("delete", null);
		templateName = parms.getParameter("TemplateName", null);		
		//from spreadsheet or textfile?
		
		
        try {
			driveService = getDriveService(securityFilePath, applicationName);
			if (deleteFile != null) {
				boolean success = deleteSheet(driveService, deleteFile);
				
		        final XQMessage message = _envelope.getMessage();
		        try {
					message.setBooleanHeader("File Deleted", success);
				} catch (XQMessageException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		        
		        _ctx.addIncomingToOutbox();

				return;
			}

	        slidesService = getSlideService(securityFilePath, applicationName);
	        sheetsService = getSheetsService(securityFilePath, applicationName);
	        if (templateName != null) {
	        	// create new presentation based on this template
	        	List<Integer> slideIndices = Arrays.asList(0, 1, 2);
	        	presentationId = createPresentationFromTemplate(slidesService,templateName, slideIndices );
	        	movePresentation(driveService, slidesService,presentationId, folderId);
	            _ctx.addIncomingToOutbox();
	            return;
	        }
	        presentationId = createPresentation(driveService, folderId, presentationName);
	        createSlides(slidesService, sheetsService, presentationName, presentationId, applicationName, dataType, dataLocation, sheetRange);
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
    
    public static void changeSheetOwner(Drive driveService, String sheetId, String newOwnerEmail) {
        try {
            // Get the current permissions on the sheet
            PermissionList permissions = driveService.permissions().list(sheetId)
                    .setFields("permissions(id, emailAddress, role)")
                    .execute();

            // Find the permission for the current owner
            String currentOwnerId = null;
            for (Permission permission : permissions.getPermissions()) {
                if (permission.getRole().equals("owner")) {
                    currentOwnerId = permission.getId();
                    break;
                }
            }

            if (currentOwnerId != null) {
                // Create a new permission for the new owner
                Permission newOwnerPermission = new Permission()
                        .setType("user")
                        .setRole("owner")
                        .setEmailAddress(newOwnerEmail);

                // Transfer ownership to the new owner
                driveService.permissions().create(sheetId, newOwnerPermission)
                        .setTransferOwnership(true)
                        .execute();

                // Remove the previous owner's permission
                driveService.permissions().delete(sheetId, currentOwnerId).execute();

                System.out.println("Sheet owner changed successfully.");
            } else {
                System.out.println("Current owner not found.");
            }
        } catch (IOException e) {
            System.out.println("Error changing sheet owner: " + e.getMessage());
        }
    }



	public static Boolean deleteSheet(Drive driveService, String sheetId) {
		boolean success = true;
		try {
			// Delete the sheet file
			driveService.files().delete(sheetId).execute();
		} catch (IOException e) {
			System.out.println("Error deleting sheet: " + e.getMessage());
			success = false;
		}
		return success;
	}
   private static String createPresentation(Drive driveService, String folderId, String presentationName) throws IOException {
        File fileMetadata = new File();
        fileMetadata.setName(presentationName);
        fileMetadata.setMimeType("application/vnd.google-apps.presentation");
        //fileMetadata.setParents(Collections.singletonList(folderId));

        File file = driveService.files().create(fileMetadata)
                .setFields("id, webViewLink")
                .execute();
        String fileId = file.getId();
        //changeSheetOwner(driveService, fileId, "horowitz@aurea.com");
        		
        return fileId; 
    }
    
    private static Presentation getPresentation(Slides slidesService, String presentationId) throws IOException {
        return slidesService.presentations().get(presentationId).execute();
    }
    
    private static void createSlides(Slides slidesService, Sheets sheetsService, String presentationName, String presentationId, String applicationName, int dataType, String dataLocation, String sheetRange) throws IOException, GeneralSecurityException { 	
    	
        Presentation presentation = getPresentation(slidesService, presentationId);
        //Create Cover Page
        createMasterSlide(slidesService, presentationId, "This Presentation is about the Product Hub");
        //addSlideWithText(slidesService, presentationId, "Cover Page", Collections.emptyList(), true );

        // Is data from a text file or from a spreadsheet
        
        switch (dataType) {
        case 0:
            List<String> slideData = readDataFromFile(dataLocation);

            // Create slides based on the data
            for (String data : slideData) {
                Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                        .setObjectId(presentationId)
                        .setInsertionIndex(presentation.getSlides().size())
                        .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TitleAndBody")));
                BatchUpdatePresentationResponse response  = slidesService.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                        .execute();

            }
        	break;
        case 1:
        	//Slides slidesService, Sheets sheetsService, String spreadsheetId, String range, String presentationId, String slideTitle)
        	createSlideFromSheet(slidesService, sheetsService, dataLocation, sheetRange, presentationId, "slideTitle");
        	break;
        default:
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
            //addSlideWithText(slidesService, presentationId, slide1Title,slide0Content, false );
            //addSlideWithText(slidesService, presentationId, slide2Title,slide1Content, false );
            //addSlideWithText(slidesService, presentationId, slide3Title,slide2Content, false );
            //addSlideWithText(slidesService, presentationId, slide4Title,slide3Content, false );
        	
        		
        }
 
 
               
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

//	    Permission permission = new Permission()
//	    		.setType("user")
//	    		.setRole("owner")
//	    		.setEmailAddress("horowitz@aurea.com");
//	    driveService.permissions().create(presentationId, permission)
//	    		.setTransferOwnership(true)
//	    		.execute();
	    
	    // Return the webViewLink of the moved presentation
	    String webViewLink = file.getWebViewLink();
	    return webViewLink;
	}

    private static  Slides getSlideService(String jsonPath, String applicationName) throws GeneralSecurityException, IOException {
		//final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

		JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();

		// Load credentials
		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

		// Load service account key from JSON file
		// Load service account key from JSON file
		GoogleCredentials credentials = GoogleCredentials.fromStream(new FileInputStream(jsonPath))
		    .createScoped(Collections.singletonList("https://www.googleapis.com/auth/drive"));

		// Build a new authorized API client service

    	
		Slides slideService = new Slides.Builder(
	            httpTransport,
	            jsonFactory,
	            new HttpCredentialsAdapter(credentials))
	        .setApplicationName(applicationName)
	        .build();  	
		return slideService;
    	
    }
	private static Drive getDriveService(String jsonPath, String applicationName)
			throws IOException, GeneralSecurityException {
		JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();

		// Load credentials
		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();

		// Load service account key from JSON file
		GoogleCredentials credentials = GoogleCredentials.fromStream(new FileInputStream(jsonPath))
			    .createScoped(Collections.singletonList("https://www.googleapis.com/auth/drive"));

	    Drive driveService = new Drive.Builder(
	            httpTransport,
	            jsonFactory,
	            new HttpCredentialsAdapter(credentials))
	        .setApplicationName(applicationName)
	        .build();


		return driveService;
	}
  
	private static Sheets getSheetsService(String jsonPath, String applicationName) throws IOException, GeneralSecurityException {

		final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

		HttpTransport httpTransport = GoogleNetHttpTransport.newTrustedTransport();
		GoogleCredentials credentials = ServiceAccountCredentials.fromStream(new FileInputStream(jsonPath))
				.createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));

		return new Sheets.Builder(httpTransport, JSON_FACTORY, new HttpCredentialsAdapter(credentials))
				.setApplicationName(applicationName).build();

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
    
    private static void addSlideWithText(Slides slidesService, String presentationId, String slideTitle, List<String> textBoxContent, boolean isCoverPage) throws IOException {
        // Create a new slide
        Presentation presentation = slidesService.presentations().get(presentationId).execute();
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout(isCoverPage ? "TITLE" : "BLANK")));

        BatchUpdatePresentationResponse response = slidesService.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest)))
                .execute();

        // Get the ID of the newly created slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();

        List<Request> requests = new ArrayList<>();

        if (isCoverPage) {
            slideId = "cover_" + System.currentTimeMillis();  // Unique ID for the slide
            String titleBoxId = "titleBox_" + System.currentTimeMillis();  // Unique ID for the text box

            // Create a cover slide with a predefined 'TITLE' layout
            Request createSlideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setObjectId(slideId)
                .setInsertionIndex(0)  // assuming it's the first slide
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("TITLE")));

            // Create a text box for the title on the cover page
            requests.add(createSlideRequest);
            requests.add(new Request().setCreateShape(new CreateShapeRequest()
                .setObjectId(titleBoxId)
                .setShapeType("TEXT_BOX")
                .setElementProperties(new PageElementProperties()
                    .setPageObjectId(slideId)
                    .setSize(new Size().setWidth(new Dimension().setMagnitude(600.0).setUnit("PT"))
                                     .setHeight(new Dimension().setMagnitude(100.0).setUnit("PT")))
                    .setTransform(new AffineTransform().setScaleX(1.0).setScaleY(1.0)
                                                        .setTranslateX(100.0).setTranslateY(50.0).setUnit("PT")))));

            // Add title text to the text box
            requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(titleBoxId)
                .setInsertionIndex(0)
                .setText(slideTitle)));

            // Update the style of the title text
            requests.add(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                .setObjectId(titleBoxId)
                .setTextRange(new Range().setType("FIXED_RANGE").setStartIndex(0).setEndIndex(slideTitle.length()))
                .setStyle(new TextStyle().setBold(true)
                                         .setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT"))
                                         .setForegroundColor(new OptionalColor().setOpaqueColor(new OpaqueColor().setRgbColor(new RgbColor().setBlue((float) 1.0).setGreen((float) 0.5).setRed((float) 0.0)))))
                .setFields("bold,fontSize,foregroundColor")));  // specify which style fields to update

            // Execute all requests in a batch update
            BatchUpdatePresentationRequest body = new BatchUpdatePresentationRequest().setRequests(requests);
            BatchUpdatePresentationResponse response2 = slidesService.presentations().batchUpdate(presentationId, body).execute();

        } else {
            // Define unique IDs for the title and text box to ensure no conflicts
            String titleShapeId = "titleShape_" + System.currentTimeMillis();
            String textBoxShapeId = "textBoxShape_" + System.currentTimeMillis();

            // Create a new shape for the title
            CreateShapeRequest titleShapeRequest = new CreateShapeRequest()
                .setObjectId(titleShapeId)
                .setShapeType("TEXT_BOX")
                .setElementProperties(new PageElementProperties()
                    .setPageObjectId(slideId)
                    .setSize(new Size()
                        .setWidth(new Dimension().setMagnitude(500.0).setUnit("PT"))
                        .setHeight(new Dimension().setMagnitude(50.0).setUnit("PT")))
                    .setTransform(new AffineTransform()
                        .setScaleX(1.0)
                        .setScaleY(1.0)
                        .setTranslateX(50.0)
                        .setTranslateY(50.0)
                        .setUnit("PT")));
            requests.add(new Request().setCreateShape(titleShapeRequest));

            // Create a new text box for the content
            CreateShapeRequest textBoxShapeRequest = new CreateShapeRequest()
                .setObjectId(textBoxShapeId)
                .setShapeType("TEXT_BOX")
                .setElementProperties(new PageElementProperties()
                    .setPageObjectId(slideId)
                    .setSize(new Size()
                        .setWidth(new Dimension().setMagnitude(500.0).setUnit("PT"))
                        .setHeight(new Dimension().setMagnitude(200.0).setUnit("PT")))
                    .setTransform(new AffineTransform()
                        .setScaleX(1.0)
                        .setScaleY(1.0)
                        .setTranslateX(50.0)
                        .setTranslateY(110.0)
                        .setUnit("PT")));
            requests.add(new Request().setCreateShape(textBoxShapeRequest));

            // Add text to the title and content text boxes in a single batch update
            requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(titleShapeId)
                .setInsertionIndex(0)
                .setText(slideTitle)));

            requests.add(new Request().setInsertText(new InsertTextRequest()
                .setObjectId(textBoxShapeId)
                .setInsertionIndex(0)
                .setText(buildTextBoxContent(textBoxContent))));

            // Execute the batch update to create the shapes and add text
            slidesService.presentations().batchUpdate(presentationId, new BatchUpdatePresentationRequest().setRequests(requests)).execute();
        }
    }

        // Helper method to build text content with formatting
 private static String buildTextBoxContent(List<String> textBoxContent) {
            StringBuilder textBoxBuilder = new StringBuilder();
            for (int i = 0; i < textBoxContent.size(); i++) {
                String line = textBoxContent.get(i);
                if (i == 0) {
                    // First line with larger font size (handled by styling if needed)
                    textBoxBuilder.append(line).append("\n");
                } else {
                    // Subsequent lines with bullets and indentation
                    textBoxBuilder.append("• ").append(line).append("\n");
                }
            }
            return textBoxBuilder.toString();
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
  
    private static void createSlideFromSheetOrg(Slides slidesService, Sheets sheetsService, String spreadsheetId, String range, String presentationId, String slideTitle)
            throws IOException, GeneralSecurityException {

 

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
        Presentation presentation = slidesService.presentations().get(presentationId).execute();
        presentation.setTitle("From Sheet");
        
        Request slideRequest = new Request().setCreateSlide(new CreateSlideRequest()
                .setInsertionIndex(presentation.getSlides().size())
                .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("BLANK")));

        
        BatchUpdatePresentationResponse slideResponse = slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(slideRequest))).execute();

        // Get the ID of the newly created slide
        String slideId = slideResponse.getReplies().get(0).getCreateSlide().getObjectId();

        // Create a new shape for the title
        List<Request> requests = new ArrayList<>();
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
        BatchUpdatePresentationResponse shapeResponse = slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        // Get the object IDs of the created shapes
        String titleId = shapeResponse.getReplies().get(0).getCreateShape().getObjectId();
        String textBoxId = shapeResponse.getReplies().get(1).getCreateShape().getObjectId();

        // Add title to the slide
        try {
            Request titleRequest = new Request().setInsertText(new InsertTextRequest()
                    .setObjectId(titleId)
                    .setInsertionIndex(0)
                    .setText(slideTitle));
            slidesService.presentations().batchUpdate(presentationId,
                    new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(titleRequest))).execute();

            Request titleStyleRequest = new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
                    .setObjectId(titleId)
                    .setTextRange(new Range()
                            .setType("FIXED_RANGE")
                            .setStartIndex(0)
                            .setEndIndex(slideTitle.length()))
                    .setStyle(new TextStyle().setFontSize(new Dimension().setMagnitude(36.0).setUnit("PT")))
                    .setFields("*"));
            slidesService.presentations().batchUpdate(presentationId,
                    new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(titleStyleRequest))).execute();
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
            slidesService.presentations().batchUpdate(presentationId,
                    new BatchUpdatePresentationRequest().setRequests(Collections.singletonList(textBoxRequest))).execute();
        } catch (GoogleJsonResponseException e) {
            // Ignore the exception if the object doesn't allow text editing
            if (!e.getDetails().getMessage().contains("does not allow text editing")) {
                throw e;
            }
        }
    }
    
    private static void createSlideFromSheet(Slides slidesService, Sheets sheetsService, String spreadsheetId, String range, String presentationId, String slideTitle)
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

    public static void createMasterSlide(Slides slidesService, String presentationId, String title) throws IOException {
        // Retrieve the presentation
        //Presentation presentation = slidesService.presentations().get(presentationId).execute();

        // Create a new slide
        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
                .setCreateSlide(new CreateSlideRequest()
                        .setObjectId(presentationId)
                        .setInsertionIndex(1)
                        .setSlideLayoutReference(new LayoutReference().setPredefinedLayout("BLANK"))));

        // Execute the request to create the slide
        BatchUpdatePresentationResponse response = slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        // Get the ID of the newly created slide
        String slideId = response.getReplies().get(0).getCreateSlide().getObjectId();

        // Prepare the requests for adding elements to the slide
        requests.clear();

        // Add the white band at the top
        requests.add(new Request()
                .setCreateShape(new CreateShapeRequest()
                        .setShapeType("RECTANGLE")
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(648.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(36.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(0.0)
                                        .setTranslateY(0.0)
                                        .setUnit("PT")))));

        // Add the logo image
        requests.add(new Request()
                .setCreateImage(new CreateImageRequest()
                		.setObjectId("LOGO_IMAGE")
                        .setUrl(LOGO_IMAGE_URL)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(72.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(36.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(576.0)
                                        .setTranslateY(0.0)
                                        .setUnit("PT")))));

        // Add the light gray background
        requests.add(new Request()
                .setCreateShape(new CreateShapeRequest()
                        .setShapeType("RECTANGLE")
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(648.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(468.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(0.0)
                                        .setTranslateY(36.0)
                                        .setUnit("PT")))));

        // Add the right image
        requests.add(new Request()
                .setCreateImage(new CreateImageRequest()
                		.setObjectId("RIGHT_IMAGE")
                        .setUrl(RIGHT_IMAGE_URL)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(144.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(144.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(504.0)
                                        .setTranslateY(144.0)
                                        .setUnit("PT")))));

        // Add the top right image
        requests.add(new Request()
                .setCreateImage(new CreateImageRequest()
                		.setObjectId("TOP_RIGHT")
                        .setUrl(TOP_RIGHT_IMAGE_URL)
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(72.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(72.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(576.0)
                                        .setTranslateY(72.0)
                                        .setUnit("PT")))));

        // Add the presentation title
        String titleShapeId = "TitleShapeID";
        requests.add(new Request()
                .setCreateShape(new CreateShapeRequest()
                		.setObjectId(titleShapeId)
                        .setShapeType("TEXT_BOX")
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(360.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(36.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(144.0)
                                        .setTranslateY(144.0)
                                        .setUnit("PT")))));

        // Add the date at the bottom
        String dateShapeId = "DateShapeID";
        requests.add(new Request()
                .setCreateShape(new CreateShapeRequest()
                		.setObjectId(dateShapeId)
                        .setShapeType("TEXT_BOX")
                        .setElementProperties(new PageElementProperties()
                                .setPageObjectId(slideId)
                                .setSize(new Size()
                                        .setWidth(new Dimension().setMagnitude(144.0).setUnit("PT"))
                                        .setHeight(new Dimension().setMagnitude(18.0).setUnit("PT")))
                                .setTransform(new AffineTransform()
                                        .setScaleX(1.0)
                                        .setScaleY(1.0)
                                        .setTranslateX(36.0)
                                        .setTranslateY(486.0)
                                        .setUnit("PT")))));
        
     // Execute the batch update to create the shapes
        BatchUpdatePresentationResponse shapeResponse = slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        // Get the object IDs of the created shapes
       // String titleShapeId = shapeResponse.getReplies().get(2).getCreateShape().getObjectId();
        //String dateShapeId = shapeResponse.getReplies().get(3).getCreateShape().getObjectId();
        requests.clear();
        
        requests.add(new Request()
                .setInsertText(new InsertTextRequest()
                        .setObjectId(titleShapeId)
                        .setInsertionIndex(0)
                        .setText(title)));
        requests.add(new Request()
                .setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(titleShapeId)
                        .setTextRange(new Range()
                                .setType("ALL"))
                        .setStyle(new TextStyle()
                                .setFontSize(new Dimension().setMagnitude(24.0).setUnit("PT"))
                                .setBold(true))
                        .setFields("*")));
        
        requests.add(new Request()
                .setInsertText(new InsertTextRequest()
                        .setObjectId(dateShapeId)
                        .setInsertionIndex(0)
                        .setText(java.time.LocalDate.now().toString())));
        
        requests.add(new Request()
                .setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(dateShapeId)
                        .setTextRange(new Range()
                                .setType("ALL"))
                        .setStyle(new TextStyle()
                                .setForegroundColor(new OptionalColor()
                                        .setOpaqueColor(new OpaqueColor()
                                                .setRgbColor(new RgbColor()
                                                        .setRed(0.2f)
                                                        .setGreen(0.2f)
                                                        .setBlue(0.2f))))
                                .setFontSize(new Dimension().setMagnitude(12.0).setUnit("PT")))
                        .setFields("*")));

        // Execute the requests to add elements to the slide
        slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }
    public static String createPresentationFromTemplate(Slides slidesService, String templatePresentationId, List<Integer> slideIndices) throws IOException {
        // Create a new presentation
        Presentation newPresentation = new Presentation()
                .setTitle("New Presentation");
        newPresentation = slidesService.presentations().create(newPresentation).execute();
        String newPresentationId = newPresentation.getPresentationId();

        // Retrieve the template presentation
        Presentation templatePresentation = slidesService.presentations().get(templatePresentationId).execute();

        // Create requests to copy the specified slides from the template presentation
        List<Request> requests = new ArrayList<>();
        for (Integer slideIndex : slideIndices) {
            String slideObjectId = templatePresentation.getSlides().get(slideIndex).getObjectId();
            if (slideObjectId.length() < 5) continue;
            requests.add(new Request()
                    .setCreateSlide(new CreateSlideRequest()
                            .setObjectId(slideObjectId)
                            .setInsertionIndex(requests.size())));
        }

        // Execute the batch update to copy the slides to the new presentation
        slidesService.presentations().batchUpdate(newPresentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();

        return newPresentationId;
    }
    
    public static void addTextToTextBox(Slides slidesService, String presentationId, String textBoxObjectId, String text) throws IOException {
        List<Request> requests = new ArrayList<>();

        // Add the text to the text box
        requests.add(new Request()
                .setInsertText(new InsertTextRequest()
                        .setObjectId(textBoxObjectId)
                        .setInsertionIndex(0)
                        .setText(text)));

        // Update the text style (optional)
        requests.add(new Request()
                .setUpdateTextStyle(new UpdateTextStyleRequest()
                        .setObjectId(textBoxObjectId)
                        .setTextRange(new Range()
                                .setType("ALL"))
                        .setStyle(new TextStyle()
                                .setFontSize(new Dimension().setMagnitude(12.0).setUnit("PT"))
                                .setBold(true)
                                .setForegroundColor(new OptionalColor()
                                        .setOpaqueColor(new OpaqueColor()
                                                .setRgbColor(new RgbColor()
                                                        .setRed((float) 0.0)
                                                        .setGreen((float) 0.0)
                                                        .setBlue((float) 0.0)))))
                        .setFields("fontSize,bold,foregroundColor")));

        // Execute the batch update
        slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }
    public static void addSlidesToPresentation(Slides slidesService, String presentationId, String templateSlideObjectId, int numberOfSlides) throws IOException {
        List<Request> requests = new ArrayList<>();

        // Retrieve the presentation
        Presentation presentation = slidesService.presentations().get(presentationId).execute();

        // Find the index of the template slide
        int templateSlideIndex = -1;
        for (int i = 0; i < presentation.getSlides().size(); i++) {
            if (presentation.getSlides().get(i).getObjectId().equals(templateSlideObjectId)) {
                templateSlideIndex = i;
                break;
            }
        }

        if (templateSlideIndex == -1) {
            throw new IllegalArgumentException("Template slide not found in the presentation.");
        }

        // Create requests to add new slides based on the template slide
        for (int i = 0; i < numberOfSlides; i++) {
            requests.add(new Request()
                    .setDuplicateObject(new DuplicateObjectRequest()
                            .setObjectId(templateSlideObjectId)
                            .setObjectIds(Collections.singletonMap(templateSlideObjectId, templateSlideObjectId + "_" + i))));

            requests.add(new Request()
                    .setUpdateSlidesPosition(new UpdateSlidesPositionRequest()
                            .setInsertionIndex(templateSlideIndex + i + 1)
                            .setSlideObjectIds(Collections.singletonList(templateSlideObjectId + "_" + i))));
        }

        // Execute the batch update
        slidesService.presentations().batchUpdate(presentationId,
                new BatchUpdatePresentationRequest().setRequests(requests)).execute();
    }
}