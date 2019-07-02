Attribute VB_Name = "mTripDistance"


Option Explicit

Function TripDistance(Origin As String, Destination As String, Optional Travel_Mode As String) As Variant
    
    '-----------------------------------------------------------------------------------------
    'The function returns the estimated trip distance (in km) between the selected origin and
    'destination using the Google Directions API.
    'NOTE: As Google points out, the use of the Google Directions API is subject to a limit
    'of 2500 requests per day, so be careful not to exceed this limit.
    
    'In order to use this function you must enable the XML, v3.0 library from VBA editor:
    'Go to Tools -> References -> check the Microsoft XML, v3.0.
    
    'Written by:    Christos Samaras
    'Date:          28/08/2013
    'e-mail:        xristos.samaras@gmail.com
    'site:          http://www.myengineeringworld.net
    '-----------------------------------------------------------------------------------------
    
    'Declaring the necessary variables. Using 30 at the first two variables because it
    'corresponds to the "Microsoft XML, v3.0" library in VBA (msxml3.dll).
    Dim Request         As New XMLHTTP30
    Dim Results         As New DOMDocument30
    Dim StatusNode      As IXMLDOMNode
    Dim DistanceNode    As IXMLDOMNode
    Dim Mode            As String
    
    'Checking if the Travel Mode value is correct (according to defined API's values).
    'Here the "transit" mode was omitted because it would require additional parameters
    '(departure time or arrival time) that would make the function quite complex.
    Select Case LCase(Travel_Mode)
        'Default, indicates standard driving directions using the road network.
        Case "driving": Mode = "driving"
        'Walking directions via pedestrian paths & sidewalks (where available).
        Case "walking": Mode = "walking"
        'Bicycling directions via bicycle paths & preferred streets (where available).
        Case "bicycling": Mode = "bicycling"
        'If driving mode was omitted, the default driving mode is used.
        Case Else: Mode = "driving"
    End Select
            
    'Check if origin is not empty.
    If Origin = "" Then
        TripDistance = "Origin is empty"
        Exit Function
    End If
    
    'Check if destination is not empty.
    If Destination = "" Then
        TripDistance = "Destination is empty"
        Exit Function
    End If
                
    'Removing possble spaces from the input variables.
    Origin = WorksheetFunction.Substitute(Origin, " ", "%20")
    Destination = WorksheetFunction.Substitute(Destination, " ", "%20")
    
    On Error GoTo errorHandler
    
    'Create the request based on Google Directions API. Parameters (from Google page):
    '- Origin:      The address or textual latitude/longitude value from which you wish to calculate directions.
                    'If you pass an address as a string, the Directions service will geocode the string and convert it to a
                    'latitude/longitude coordinate to calculate directions.
                    'If you pass coordinates, ensure that no space exists between the latitude and longitude values.
                    
    '- Destination: The address or textual latitude/longitude value from which you wish to calculate directions.
                    'If you pass an address as a string, the Directions service will geocode the string and convert it to a
                    'latitude/longitude coordinate to calculate directions.
                    'If you pass coordinates, ensure that no space exists between the latitude and longitude values.
                    
    '- Ssensor:     Indicates whether or not the directions request comes from a device with a location sensor.
                    'This value must be either true or false.
                    
    '- Mode:        Optional parameter - default value is "driving". Specifies the mode of transport to use when calculating directions.

    Request.Open "GET", "https://maps.googleapis.com/maps/api/directions/xml?" _
    & "origin= " & Origin & "&destination=" & Destination & "&sensor=false" & "&mode=" & Mode & "&key=API_KEY", False
            
    'Send the request to the Google server.
    Request.send
    
    'Read the results from the request.
    Results.LoadXML Request.responseText
    
    'Get the status node value.
    Set StatusNode = Results.SelectSingleNode("//status")
    
    'Based on the status node result, proceed accordingly.
    Select Case StatusNode.Text
            
        Case "OK"   'The response contains a valid result.
            'Get the distnance node value and calculate trip distance in km.
            Set DistanceNode = Results.SelectSingleNode("//leg/distance/value")
            TripDistance = CDbl(DistanceNode.Text) / 1000
    
        Case "INVALID_REQUEST"  'The provided request was invalid.
                                'Common causes of this status include an invalid parameter or parameter value.
            TripDistance = "Invalid request"
        
        Case "NOT_FOUND"    'At least one of the locations specified in the requests's origin,
                            'destination, or waypoints could not be geocoded.
            TripDistance = "Origin/destination could not be geocoded"
                    
        Case "ZERO_RESULTS" 'No route could be found between the origin and destination.
            TripDistance = "Could not find route"
                                
        Case "MAX_WAYPOINTS_EXCEEDED"   'Too many waypoints were provided in the request The maximum allowed waypoints is 8, plus the origin, and destination.
                                        '(Google Maps API for Business customers may contain requests with up to 23 waypoints.)
            TripDistance = "Too many waypoints"
            
        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded limit.
            TripDistance = "Requestor has exceeded limit"
            
        Case "REQUEST_DENIED"   'The service denied use of the directions service.
            TripDistance = "Invalid sensor parameter"
        
        Case "UNKNOWN_ERROR"    'The request could not be processed due to a server error.
            TripDistance = "Server error"
        
        Case Else   'Just in case...
            TripDistance = "Error"
        
    End Select
        
    'In case of error, release the objects.
errorHandler:
    Set StatusNode = Nothing
    Set DistanceNode = Nothing
    Set Results = Nothing
    Set Request = Nothing
    
End Function




