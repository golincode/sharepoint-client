# SharePoint REST API Client

The **SharePoint REST API Client** allows you to work with a subset of SharePoint's functionality (Lists and Users for now) using PHP.

The library is available to anyone and is licensed under the MIT license.

### Known to work with the following versions:
* SharePoint Online (2013)

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

Class instantiation:

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

In order to use the **SharePoint REST API Client** methods, you need an **Access Token** which can be generated either by a logged User or an Application.

Get an Access Token as a **User**:

    try {
        $config = array(
            'token'  => 'eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOiIyNTQyNGR...', // Context Token
            'secret' => 'YzcZQ7N4lTeK5COin/nmNRG5kkL35gAW1scrum5mXVgE='
        );

        $rc->tokenFromUser($config);

        /**
         * From this point on, we have an Access Token
         * which allows us to use the other class methods
         */

    } catch(SharePointException $e) {
        // handle exceptions
    }


Get an Access Token as an **Application**:

    try {
        $config = array(
            'acs'       => 'https://accounts.accesscontrol.windows.net/tokens/OAuth/2', // Access Control Service URL
            'client_id' => '52848cad-bc13-4d69-a371-30deff17bb4d/example.com@09g7c3b0-f0d4-416d-39a7-09671ab91f64',
            'secret'    => 'YzcZQ7N4lTeK5COin/nmNRG5kkL35gAW1scrum5mXVgE=',
            'resource'  => '00000000-0000-ffff-0000-000000000000/example-69z86c039b91sn.sharepoint.com@09g7c3b0-f0d4-416d-39a7-09671ab91f64'
        );

        $rc->tokenFromUser($config);

        /**
         * From this point on, we have an Access Token
         * which allows us to use the other class methods
         */

    } catch(SharePointException $e) {
        // handle exceptions
    }

Both methods return an Array.
