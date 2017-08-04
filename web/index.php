<?php



echo '
<html >

  <body>
 
<div class="pen-title">
  <h1>CREATE EXCEL INVOICE</h1>
</div>
<!-- Form Module-->
<div class="module form-module">
  <div class="toggle"><i class="fa fa-times fa-pencil"></i>
    
  </div>
  <div class="form">
    <h2>TYPE ORDER NUMBER HERE</h2>
    <fieldset>
    <form method="post" action="invoice.php" style="alignment-adjust:central"> 
                    <label for="xl_order_id">Order#:</label><br/>
                    <input type="text" name="xl_order_id" placeholder="365"/>
  
      <button>CREATE INVOICE</button>
    </form>
    </fieldset> 
  </div>
 
    
  </body>
</html>

';

?>
