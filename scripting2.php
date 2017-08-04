<?php
use Bigcommerce\Api\Client as Bigcommerce;
//connect to SilverForte
 Bigcommerce::configure(array(

        'store_url' => 'https://silverforte.com',
        'username' => 'scripting2',
        'api_key' => '736d29852c9c69fbf7e6f84682e9e2e33577679e'
    ));

Bigcommerce::verifyPeer(false);

?>