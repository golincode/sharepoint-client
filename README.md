# SharePoint REST API Client

The **SharePoint REST API Client** allows you to authenticate with SharePoint Online (2013) via [OAuth2](http://oauth.net/2/) and work with a subset of SharePoint's functionality (currently **Lists** and **Users**) using [PHP](http://www.php.net).

The library is available to anyone and is licensed under the MIT license.


## Installation

### Requirements
* PHP 5.4+
* [Guzzle](https://packagist.org/packages/guzzlehttp/guzzle)
* [PHP-JWT](https://packagist.org/packages/nixilla/php-jwt)
* [Carbon](https://packagist.org/packages/nesbot/carbon)


#### Using Composer

Add [altek/sharepoint-client](https://packagist.org/packages/altek/sharepoint-client) to your `composer.json` and run **composer install** or **composer update**.

    {
        "require": {
            "altek/sharepoint-client": "*"
        }
    }


## Usage

### Class instantiation:

    <?php

    require 'vendor/autoload.php';

    use Altek\SharePoint\SharePointException;
    use Altek\SharePoint\RESTClient;

    try {
        $config = array(
            'url'  => 'https://example-69z86c039b91sn.sharepoint.com', // Application URL
            'path' => '/sites/mySite'                                  // Application URL Path
        );

        $rc = new RESTClient($config);

    } catch(SharePointException $e) {
        // handle exceptions
    }

**Attention:** In order to use the methods provided by this class, you need an **Access Token** which can be requested either by a logged User or an **Application only policy**.


### Get an Access Token as a **User**:

    try {
        $config = array(
            'token'  => 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIyNTQyNGR...', // Context Token
            'secret' => 'YzcZQ7N4lTeK5COin/nmNRG5kkL35gAW1scrum5mXVgE='
        );

        $rc->tokenFromUser($config);

    } catch(SharePointException $e) {
        // handle exceptions
    }

**Note:** The context token is accessible via the **SPAppToken** input (POST) when SharePoint redirects to our Application.


### Get an Access Token as an **Application**:

    try {
        $config = array(
            'acs'       => 'https://accounts.accesscontrol.windows.net/tokens/OAuth/2', // Access Control Service URL
            'client_id' => '52848cad-bc13-4d69-a371-30deff17bb4d/example.com@09g7c3b0-f0d4-416d-39a7-09671ab91f64',
            'secret'    => 'YzcZQ7N4lTeK5COin/nmNRG5kkL35gAW1scrum5mXVgE=',
            'resource'  => '00000000-0000-ffff-0000-000000000000/example-69z86c039b91sn.sharepoint.com@09g7c3b0-f0d4-416d-39a7-09671ab91f64'
        );

        $rc->tokenFromApp($config);

    } catch(SharePointException $e) {
        // handle exceptions
    }

Both methods return an Array with the access token and start/expire dates as [Carbon](https://packagist.org/packages/nesbot/carbon) objects.
The access token will be stored internally so that all methods that require it, have access to it.


### Get the Context Web Information

    try {
        $ctx = $rc->getContextInfo();

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with the SharePoint's library version and the Form Digest (needed to update, upload and delete Items).


## List methods

### Get available Lists

    try {
        // Extra properties to add
        $extra = array(
            'is_hidden' => 'Hidden',
            // ...
        );

        $lists = $rc->getLists($extra);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with all the available Lists.

**Note:** By default, a List Array will only contain: **GUID**, **Title**, **Description**, **Items**, **Created** date as a [Carbon](https://packagist.org/packages/nesbot/carbon) object.

If more information about a List is needed, use the **$extra** Array parameter to add other properties returned by the SharePoint API to the resulting Array.


### Get the Item count of a List

    try {
        $library = 'MyLibrary';

        $count = $rc->getListItemCount($library);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an integer with the number of Items that exist in a List.


### Get the Items from a specific List

    try {
        $library = 'MyLibrary';

        // Extra properties to add
        $extra = array(
            'content_type_id' => 'ContentTypeId',
            // ...
        );

        $items = $rc->getListItems($library, $extra);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with all the available Items from a List.

**Note:** By default, an Item Array will only contain: **ID**, **GUID**, **Title**, **Entity Type**, **URI**, **Author ID**, **Editor ID**, and **Created/Modified** date as [Carbon](https://packagist.org/packages/nesbot/carbon) objects.

If more information about an Item is needed, use the **$extra** Array parameter to add other properties returned by the SharePoint API to the resulting Array.


### Get an Item from a specific List

    try {
        $library = 'MyLibrary';

        $id = 1;

        // Extra properties to add
        $extra = array(
            'content_type_id' => 'ContentTypeId',
            // ...
        );

        $item = $rc->getListItem($library, $id, $extra);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with the Item information.

**Note:** By default, an Item Array will only contain: **ID**, **GUID**, **Title**, **Entity Type**, **URI**, **Author ID**, **Editor ID**, and **Created/Modified** date as [Carbon](https://packagist.org/packages/nesbot/carbon) objects.

If more information about an Item is needed, use the **$extra** Array parameter to add other properties returned by the SharePoint API to the resulting Array.


### Upload an Item to a List

    try {
        $library = 'MyLibrary';

        $file = '/path/to/a/picture.jpg';

        // Item properties
        $properties = array(
            'Title' => 'A picture in JPEG format',
            // ...
        );

        // Silently overwrite an existing file with the same name
        $overwrite = true;

        $rc->uploadListItem($library, $file, $properties, $overwrite);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns true if the file was successfully uploaded.

**Note:** If the **$overwrite** value is false, an exception will be thrown when uploading a file with a name that already exists in the List.


### Update a List Item

    try {
        $library = 'MyLibrary';

        $id = 1;

        // Item properties
        $properties = array(
            'Title' => 'A different title for the picture in JPEG format',
            // ...
        );

        // Entity Type
        $type = 'SP.Data.ReportsListItem';

        $rc->updateListItem($library, $id, $properties, $type);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns true if the file was successfully updated.

**Note:** The Entity Type of an Item can be fetched using the **getListItem()** method.


### Delete a List Item

    try {
        $library = 'MyLibrary';

        $id = 1;

        $rc->deleteListItem($library, $id);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns true if the file was successfully deleted.

**Note:** The Entity Type of an Item can be fetched using the **getListItem()** method.


## User methods

### Get the current User profile

    try {
        $user = $rc->getCurrentUserProfile();

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with the User information.

**Note:** The User information Array will only contain: **Account Name**, **Email**, **Name**, **URL**, **Picture URL** and **Title**.

**Attention:** This method will throw an exception if the Access Token was generated by an **Application only policy**.


### Get a specific User profile

    try {
        $account = 'i:0#.f|membership|username@example.onmicrosoft.com';

        $user = $rc->getUserProfile($account);

    } catch(SharePointException $e) {
        // handle exceptions
    }

This method returns an Array with the User information.

**Note:** The User information Array will only contain: **Account Name**, **Email**, **Name**, **URL**, **Picture URL** and **Title**.


## Extra List/Item properties

Note that extra properties of **date** or **datetime** types, will be returned as [Carbon](https://packagist.org/packages/nesbot/carbon) objects.
