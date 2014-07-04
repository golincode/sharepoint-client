<?php
/**
 * This file is part of the SharePoint REST API Client package.
 *
 * @author     Quetzy Garcia <quetzyg@altek.org>
 * @copyright  2014 Quetzy Garcia
 *
 * For the full copyright and license information,
 * please view the LICENSE file that was distributed
 * with this source code.
 */


namespace Altek\SharePoint;

use Carbon\Carbon;
use GuzzleHttp\Client;
use GuzzleHttp\Exception\ParseException;
use GuzzleHttp\Exception\RequestException;
use JWT\Authentication\JWT;
use UnexpectedValueException;


/**
 * SharePoint REST API Client
 */
class RESTClient
{
	/**
	 * REST Client object
	 *
	 * @access  private
	 */
	private $client = null;


	/**
	 * Application URL path
	 *
	 * @access  private
	 */
	private $path = null;


	/**
	 * Access Token
	 *
	 * @access  private
	 */
	private $access_token = null;


	/**
	 * Token source (User or App?)
	 *
	 * @access  private
	 */
	private $from_app = true;


	/**
	 * REST Client constructor
	 *
	 * @access  public
	 * @param   array $config
	 * @throws  SharePointException
	 * @return  RESTClient
	 */
	public function  __construct(array $config)
	{
		if (empty($config['url'])) {
			throw new SharePointException('The Application URL is empty/not set');
		}

		if ( ! filter_var($config['url'], FILTER_VALIDATE_URL)) {
			throw new SharePointException('The Application URL is invalid');
		}

		if (empty($config['path'])) {
			throw new SharePointException('The Application URL path is empty/not set');
		}

		$this->path = $config['path'];

		$this->client = new Client(array(
			'base_url' => $config['url']
		));

		/**
		 * Set default cURL options
		 */
		$this->client->setDefaultOption('config', array(
			'curl' => array(
				CURLOPT_SSLVERSION     => 3,
				CURLOPT_SSL_VERIFYHOST => 0,
				CURLOPT_SSL_VERIFYPEER => 0
			)
		));
	}


	/**
	 * Free resources
	 *
	 * @access  public
	 * @return  void
	 */
	public function __destruct()
	{
		$this->client = null;
	}


	/**
	 * Add object properties to an array
	 *
	 * @access  private
	 * @param   object $obj        Object to map from
	 * @param   array  $properties Array to add properties to
	 * @param   array  $map        Key to object properties
	 * @throws  SharePointException
	 * @return  void
	 */
	private function addProperties($obj = null, array &$properties, array $map = array())
	{
		foreach($map as $key => $property) {
			$property = str_replace(' ', '_x0020_', $property);

			if ( ! property_exists($obj, $property)) {
				throw new SharePointException('Invalid property: '.$property);
			}

			// remove metadata
			if ($obj->$property instanceof \stdClass && property_exists($obj->$property, '__metadata')) {
				unset($obj->$property->__metadata);
			}

			// create a Carbon object if a date/time is detected
			if (is_string($obj->$property) && preg_match('/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/', $obj->$property) === 1) {
				$obj->$property = new Carbon($obj->$property);
			}

			$properties[$key] = $obj->$property;
		}
	}


	/**
	 * Get Item properties
	 *
	 * @access  private
	 * @param   object $item  Item object
	 * @param   array  $extra Extra properties to add to the array
	 * @throws  SharePointException
	 * @return  array
	 */
	private function getItemProperties($item = null, array $extra = array())
	{
		// default properties
		$properties = array(
			'id'          => $item->Id,
			'guid'        => $item->GUID,
			'title'       => $item->Title,
			'entity_type' => $item->__metadata->type,
			'uri'         => $item->__metadata->uri,
			'created'     => new Carbon($item->Created),
			'modified'    => new Carbon($item->Modified),
			'author_id'   => $item->AuthorId,
			'editor_id'   => $item->EditorId
		);

		// add extra properties
		$this->addProperties($item, $properties, $extra);

		return $properties;
	}


	/**
	 * Get an Authentication Token through a User
	 *
	 * @access  public
	 * @param   array $config
	 * @throws  SharePointException
	 * @return  array
	 */
	public function tokenFromUser(array $config)
	{
		if (empty($config['token'])) {
			throw new SharePointException('The Context Token is empty/not set');
		}

		if (empty($config['secret'])) {
			throw new SharePointException('The Secret is empty/not set');
		}

		try {
			$jwt = JWT::decode($config['token'], $config['secret'], false);

			// get URL hostname
			$hostname = parse_url($this->client->getBaseUrl(), PHP_URL_HOST);

			// build resource
			$resource = str_replace('@', '/'.$hostname.'@', $jwt->appctxsender);

			// decode App context
			$oauth2 = json_decode($jwt->appctx);

			$response = $this->client->post($oauth2->SecurityTokenServiceUri, array(
				'exceptions' => false,

				'headers'    => array(
					'Content-Type' => 'application/x-www-form-urlencoded'
				),

				/**
				 * The POST data has to be passed as a query string
				 */
				'body'       => http_build_query(array(
					'grant_type'    => 'refresh_token',
					'client_id'     => $jwt->aud,
					'client_secret' => $config['secret'],
					'refresh_token' => $jwt->refreshtoken,
					'resource'      => $resource
				))
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			$this->access_token = $json->access_token;

			$this->from_app = false;

			return array(
				'access_token' => $json->access_token,
				'not_before'   => Carbon::createFromTimestamp($json->not_before),
				'expires_on'   => Carbon::createFromTimestamp($json->expires_on)
			);

		// Guzzle
		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());

		// JWT
		} catch(UnexpectedValueException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get an Authentication Token through an Application
	 *
	 * @access  public
	 * @param   array $config
	 * @throws  SharePointException
	 * @return  array
	 */
	public function tokenFromApp(array $config)
	{
		if (empty($config['acs'])) {
			throw new SharePointException('The Access Control Service URL is empty/not set');
		}

		if ( ! filter_var($config['acs'], FILTER_VALIDATE_URL)) {
			throw new SharePointException('The Access Control Service URL is invalid');
		}

		if (empty($config['client_id'])) {
			throw new SharePointException('The Client ID is empty/not set');
		}

		if (empty($config['secret'])) {
			throw new SharePointException('The Secret is empty/not set');
		}

		if (empty($config['resource'])) {
			throw new SharePointException('The Resource is empty/not set');
		}

		try {
			$response = $this->client->post($config['acs'], array(
				'exceptions' => false,

				'headers'    => array(
					'Content-Type' => 'application/x-www-form-urlencoded'
				),

				/**
				 * The POST data has to be passed as a query string
				 */
				'body'       => http_build_query(array(
					'grant_type'    => 'client_credentials',
					'client_id'     => $config['client_id'],
					'client_secret' => $config['secret'],
					'resource'      => $config['resource']
				))
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			$this->access_token = $json->access_token;

			$this->from_app = true;

			return array(
				'access_token' => $json->access_token,
				'not_before'   => Carbon::createFromTimestamp($json->not_before),
				'expires_on'   => Carbon::createFromTimestamp($json->expires_on)
			);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get the Context Web Information
	 *
	 * @access  private
	 * @throws  SharePointException
	 * @return  array
	 */
	public function getContextInfo()
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			$response = $this->client->post($this->path.'/_api/contextinfo', array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return array(
				'lib_version' => $json->d->GetContextWebInformation->LibraryVersion,
				'form_digest' => array(
					'value'   => $json->d->GetContextWebInformation->FormDigestValue,
					'timeout' => $json->d->GetContextWebInformation->FormDigestTimeoutSeconds,
					'expires' => time() + $json->d->GetContextWebInformation->FormDigestTimeoutSeconds
				)
			);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get all the available Lists
	 *
	 * @access  public
	 * @param   array $extra Extra properties to add to the array
	 * @throws  SharePointException
	 * @return  array
	 */
	public function getLists(array $extra = array())
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			$response = $this->client->get($this->path.'/_api/web/Lists', array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			$lists = array();

			foreach($json->d->results as $list) {
				// default properties
				$properties = array(
					'guid'        => $list->Id,
					'title'       => $list->Title,
					'description' => $list->Description,
					'items'       => $list->ItemCount,
					'created'     => new Carbon($list->Created)
				);

				// add extra properties
				$this->addProperties($list, $properties, $extra);

				$lists[] = $properties;
			}

			return $lists;

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get the List Item (Document) count
	 *
	 * @access  public
	 * @param   string $library Library name
	 * @throws  SharePointException
	 * @return  int the number of Items in the List
	 */
	public function getListItemCount($library = null)
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			$response = $this->client->get($this->path."/_api/web/Lists/GetByTitle('".$library."')/itemCount", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return $json->d->ItemCount;

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get List Items (Documents)
	 *
	 * @access  public
	 * @param   string $library Library name
	 * @param   array  $extra   Extra properties to add to the array
	 * @throws  SharePointException
	 * @return  array of arrays with Item properties
	 */
	public function getListItems($library = null, array $extra = array())
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			$response = $this->client->get($this->path."/_api/web/Lists/GetByTitle('".$library."')/items", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			$items = array();

			foreach($json->d->results as $item) {
				$items[] = $this->getItemProperties($item, $extra);
			}

			return $items;

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get a List Item (Document) by ID
	 *
	 * @access  public
	 * @param   string $library Library name
	 * @param   int    $id      Item (Document) ID
	 * @param   array  $extra   Extra properties to add to the array
	 * @throws  SharePointException
	 * @return  array with Item properties
	 */
	public function getListItem($library = null, $id = 0, array $extra = array())
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			if (empty($id)) {
				throw new SharePointException('The Item ID is empty/not set');
			}

			$response = $this->client->get($this->path."/_api/web/Lists/GetByTitle('".$library."')/items(".$id.")", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return $this->getItemProperties($json->d, $extra);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Upload a List Item (Document)
	 *
	 * @access  public
	 * @param   string $library    Library name
	 * @param   string $file       Path to the file to be uploaded
	 * @param   array  $properties Item (Document) properties to insert
	 * @param   bool   $overwrite  Allow the file to overwrite an existing one
	 * @throws  SharePointException
	 * @return  bool true if the Item was uploaded
	 */
	public function uploadListItem($library = null, $file = null, array $properties, $overwrite = false)
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			if (is_readable($file) === false) {
				throw new SharePointException('The following file does not exist or cannot be read: '.$file);
			}

			$ctx = $this->getContextInfo();

			$data = file_get_contents($file);

			if ($data === false) {
				throw new SharePointException('Failure to get the file contents for: '.$file);
			}

			$response = $this->client->post($this->path."/_api/web/GetFolderByServerRelativeUrl('Lists/".$library."')/Files/Add(url='".basename($file)."',overwrite='".($overwrite ? 'true' : 'false')."')", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization'   => 'Bearer '.$this->access_token,
					'Accept'          => 'application/json;odata=verbose',
					'X-RequestDigest' => $ctx['form_digest']['value']
				),

				'query'      => array(
					'$select' => 'ListItemAllFields/Id',
					'$expand' => 'ListItemAllFields'
				),

				'body'       => $data
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			$id = $json->d->ListItemAllFields->Id;
			$type = $json->d->ListItemAllFields->__metadata->type;

			/**
			 * Update the Item's metadata
			 */
			return $this->updateListItem($library, $id, $properties, $type);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Update an existing Item (Document)
	 *
	 * @access  public
	 * @param   string $library    Library name
	 * @param   int    $id         Item (Document) ID
	 * @param   array  $properties Item (Document) properties to update
	 * @param   string $type       List Item Entity Type
	 * @throws  SharePointException
	 * @return  bool true if the Item was updated
	 */
	public function updateListItem($library = null, $id = 0, array $properties, $type = null)
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			if (empty($id)) {
				throw new SharePointException('The Item ID is empty/not set');
			}

			if (empty($type)) {
				throw new SharePointException('The Entity Type is empty/not set');
			}

			$ctx = $this->getContextInfo();

			$post = array('__metadata' => array(
				'type' => $type
			));

			$post = array_merge($post, $properties);

			$data = json_encode($post);

			$response = $this->client->post($this->path."/_api/web/Lists/GetByTitle('".$library."')/items(".$id.")", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization'   => 'Bearer '.$this->access_token,
					'Accept'          => 'application/json;odata=verbose',
					'X-RequestDigest' => $ctx['form_digest']['value'],
					'X-HTTP-Method'   => 'MERGE',
					'IF-MATCH'        => '*', // always match the eTag
					'Content-type'    => 'application/json;odata=verbose',
					'Content-length'  => strlen($data)
				),

				'body'       => $data
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return true;

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Delete a List Item (Document)
	 *
	 * @access  public
	 * @param   string $library Library name
	 * @param   int    $id      Item (Document) ID
	 * @throws  SharePointException
	 * @return  bool true if the Item was deleted
	 */
	public function deleteListItem($library = null, $id = 0)
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($library)) {
				throw new SharePointException('The Library is empty/not set');
			}

			if (empty($id)) {
				throw new SharePointException('The Item ID is empty/not set');
			}

			$ctx = $this->getContextInfo();

			$response = $this->client->post($this->path."/_api/web/Lists/GetByTitle('".$library."')/items(".$id.")", array(
				'exceptions' => false,

				'headers'    => array(
					'Authorization'   => 'Bearer '.$this->access_token,
					'Accept'          => 'application/json;odata=verbose',
					'X-RequestDigest' => $ctx['form_digest']['value'],
					'X-HTTP-Method'   => 'DELETE',
					'IF-MATCH'        => '*' // always match the eTag
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return true;

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get the profile of the currently logged User
	 *
	 * @access  public
	 * @throws  SharePointException
	 * @return  array
	 */
	public function getCurrentUserProfile()
	{
		try {
			if ($this->from_app) {
				throw new SharePointException('This method can only be called when the Access Token originated from a User');
			}

			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() first');
			}

			$response = $this->client->get($this->path.'/_api/SP.UserProfiles.PeopleManager/GetMyProperties', array(
				'exceptions' => false,

				'headers' => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return array(
				'account' => $json->d->AccountName,
				'email'   => $json->d->Email,
				'name'    => $json->d->DisplayName,
				'url'     => $json->d->PersonalUrl,
				'picture' => $json->d->PictureUrl,
				'title'   => $json->d->Title
			);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}


	/**
	 * Get the profile of a specific User
	 *
	 * @access  public
	 * @param   string $account User account
	 * @throws  SharePointException
	 * @return  array
	 */
	public function getUserProfile($account = null)
	{
		try {
			if (empty($this->access_token)) {
				throw new SharePointException('The Access Token does not exist. Run tokenFromUser() or tokenFromApp() first');
			}

			if (empty($account)) {
				throw new SharePointException('The Account is empty/not set');
			}

			$response = $this->client->post($this->path."/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)", array(
				'exceptions' => false,

				'headers' => array(
					'Authorization' => 'Bearer '.$this->access_token,
					'Accept'        => 'application/json;odata=verbose'
				),

				'query' => array(
					'@v' => "'".$account."'"
				)
			));

			$json = $response->json(array('object' => true));

			if (isset($json->error)) {
				throw new SharePointException($json->error->message->value);
			}

			return array(
				'account' => $json->d->AccountName,
				'email'   => $json->d->Email,
				'name'    => $json->d->DisplayName,
				'url'     => $json->d->PersonalUrl,
				'picture' => $json->d->PictureUrl,
				'title'   => $json->d->Title
			);

		} catch(ParseException $e) {
			throw new SharePointException($e->getMessage());
		}
	}
}
