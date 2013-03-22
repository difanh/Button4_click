#include "stdafx.h"
#include <time.h>

// Forward declarations
static HRESULT GetInventorInformation();

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT Result = NOERROR;
	Result = ::CoInitialize (NULL);

	ApplicationPtr pInvApp;
	pInvApp.GetActiveObject("Inventor.Application");
	if(pInvApp == NULL)
	{
		printf("There is no Inventor open!\n");
		return 1;
	}

	// Check to see there is a document open.
	if (pInvApp->GetDocuments()->GetCount() == 0)
	{
		printf( "There is no document open!\n" );
		return E_FAIL;
	}

	//printf("%d\t, No Document: %d\n", pInvApp->GetActiveDocumentType(), kNoDocument);

	//Get the .ipt Document (kPartDocumentObject = 12290)
	DocumentPtr pItem;
	pItem = pInvApp->GetActiveDocument();
	if(pItem == NULL)
	{
		printf("Error: No active document!\n");
		return E_FAIL;
	}

	//Get the activated item as a part document
	PartDocumentPtr pPartDoc;
	if(pItem->GetDocumentType() == kPartDocumentObject)
		pPartDoc = pItem;
	if(pPartDoc == NULL)
	{
		printf("The active document is not a part document!\n");
		return E_FAIL;
	}
	printf("Part Document Opened successfully!\n");

	//Get the part document definition
	PartComponentDefinitionPtr pCompDef;
	pCompDef = pPartDoc->GetComponentDefinition();

	//Get the selected item and the selected face
	SelectSetPtr selectedItem;
	selectedItem = pPartDoc->GetSelectSet();
	if(selectedItem->GetCount() == 0)
	{
		printf("No items in this part document is selected!\n");
		return E_FAIL;
	}

	FacePtr face;
	//printf("\n*************Face Selection***************\n");
	face = selectedItem->GetItem(1);
	if(face->GetType() != kFaceObject)
	{
		printf("The selected item is not a plane-face!\n");
	}

	// Get the selected face as workplane
	WorkPlanePtr pWorkPlane; 
	tagVARIANT offset;
	offset.vt = sizeof(long);
	offset.lVal = 10;
	if(pPartDoc->ComponentDefinition->WorkPlanes->AddByPlaneAndOffset(face, offset, 0, &pWorkPlane) == NOERROR)
		printf("Succeeded to set the workplane!\n");
	else
	{
		printf("Fail to set Work Plane!\n");
		return E_FAIL;
	}

	//Create a sketch based on the work plane just selected and set the transient geometry
	PlanarSketchPtr pSketch;
	if(pPartDoc->ComponentDefinition->Sketches->Add(pWorkPlane, false, &pSketch) == NOERROR)
		printf("Suceeded to set the sketch!\n");

	TransientGeometryPtr pTG;
	pTG = pInvApp->GetTransientGeometry();

	//We can use a camera to rotate the object to any position, remember to apply
	CameraPtr pCamera;
	pCamera = pInvApp->ActiveView->GetCamera();
	pCamera->PutViewOrientationType(kIsoTopLeftViewOrientation);
	pCamera->Apply();

    //	pSketch->SketchPoints->Add();
	printf("**\n");
	time_t Tstart;
	Tstart = time(NULL);
	//Add a point in the sketch with x coordinate and y coordinate
	Point2dPtr point;
	int i, j;
	int m = 10;
	int n = 10;
	for(i = 0; i < m; i++)
	{
		for(j = 0; j < n; j++)
		{
			pInvApp->TransientGeometry->CreatePoint2d(i*0.3, j*0.3, &point);

			//Sketch a circle using the point defined and the radius
			SketchCirclePtr pCircle;
			pSketch->SketchCircles->AddByCenterRadius(point, 0.1, &pCircle);
			//printf("%ld\n",pSketch->SketchEntities->GetCount());		
		}
	}
	time_t Tend = time(NULL);
	int num = m*n;
	time_t t = Tend - Tstart;
	printf("Items drawn: %d\n", num);
	printf("It took %lu seconds\n", t);
	
	//Create new profile
	ProfilePtr pProfile;
	tagVARIANT profilePathSegments;
	tagVARIANT reserved;
	//Fail to create profile
	if(pSketch->Profiles->AddForSolid(true, profilePathSegments, reserved,  &pProfile) != NOERROR)
	{
		printf("Fail to create profile!(AddForSolid)\n");
		return E_FAIL;
	}
	
	//Create a new Extrution definition method
	ExtrudeDefinitionPtr pExtrudeDef;
	
	if(pCompDef->Features->ExtrudeFeatures->CreateExtrudeDefinition(pProfile, kCutOperation, &pExtrudeDef) != NOERROR)
	{
		printf("Fail to create extrude definition\n!");
		return E_FAIL;
	}

	VARIANT distance;
	distance.vt = sizeof(double);
	distance.dblVal = 1.0;
	
	if(pExtrudeDef->SetDistanceExtent(distance, kSymmetricExtentDirection) != NOERROR)
		printf("Set Extrude Extent Fail\n");
	else
		printf("Extrution Feature Set!\n");
	
	//Do the extrusion
	ExtrudeFeaturePtr pExtrude;
	pCompDef->Features->ExtrudeFeatures->Add(pExtrudeDef, &pExtrude);
	
	return 0;

}
