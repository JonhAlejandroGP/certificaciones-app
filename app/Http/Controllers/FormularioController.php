<?php
namespace App\Http\Controllers;
require_once '../vendor/autoload.php';
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWord\Style\Language;
use PhpOffice\PhpWord\Style\ListItem;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use App\Providers\DataImport;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpWord\Shared\Converter;
use PDO;
use NumberToWords\NumberToWords;

class FormularioController extends Controller
{
    private $pdo;

    public function __construct()
    {
        // $dsn = 'oci:dbname=//129.213.43.0:1521/SICAPITA';
        // $username = 'CERTIFICADO_LAB';
        // $password = 'certificado_lab';
        $dsn = 'mysql:host=localhost;dbname=wallet58';
        $username = 'root';
        $password = 'Consulta.';       
        
        try {
            // Crear una instancia de PDO
            $this->pdo = new PDO($dsn, $username, $password);
            $this->pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

        } catch (PDOException $e) {
            
            die('Error de conexión: ' . $e->getMessage());
        }
    }

    public function buscar(Request $request)
    {
        // Recibir el parámetro del formulario
        $parametro = $request->input('parametro');      
              if($parametro==0)
           {
            $sql = 'SELECT DISTINCT NUMERO_IDENTIFICACION FROM funcionarios1';
            //$sql = "SELECT DISTINCT NUMERO_IDENTIFICACION FROM rh_actos_x_funcionario"
            $stmt = $this->pdo->prepare($sql);
            $stmt->execute();
            $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);
            foreach ($resultados as $fila) {
                $this -> crear_word1($fila['NUMERO_IDENTIFICACION'],2);    
            }
            return view('Formulario', ['mensaje' => 'Certificados Generados en la ubicacion C:/certificacioneslaborales']);   
           }else{
            $this -> crear_word1($parametro,1);       
           }        
    }

    public function buscarFirma()
    {
        // Recibir el parámetro del formulario        
            $sql = "select nombres||' '||primer_apellido||' '||segundo_apellido as nombrejefe from rh_personas join rh_dependencias on interno_persona=cod_jefe and codigo_dependencia=12110";
            $stmt = $this->pdo->prepare($sql);
            $stmt->execute();
            $resultado = $stmt->fetchAll(PDO::FETCH_ASSOC);            
            $firma=$resultado['nombrejefe'];
            return $firma;
    }

    public function procesarArchivo(Request $request)
    {
        $request->validate([
            'archivo' => 'required|file|mimes:xls,xlsx'
        ]);

        $file = $request->file('archivo');
        $filePath = $file->getPathname();        
        $spreadsheet = IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray(null, true, true, true);
        $generamasivo=2;
        
        foreach ($data as $row) {
        
            $parametro = $row['A'];          
          $this ->  crear_word1($parametro,$generamasivo);
           
        }
       
    }

    public function crear_word1($parametro,$tipogenera)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $phpWord->setDefaultFontName('Arial');
        $phpWord->setDefaultFontSize(12);
        $hayresol = 0;
        $t_acto =0;
        $n_resolucion=0;
        $numero_acto = "";
        $filafuncion="";
        $mensaje;

        $section = $phpWord->addSection();    
        $header = $section->addHeader();
       
        $header->addImage(
            '../resources/persoenc.png', 
            array('wrappingStyle' => 'square',
            'width' => $this -> cmToPt(21.59),  
            'height' => $this -> cmToPt(3),
            'marginTop' => 0,
            'marginLeft' => -50,
            'wrapDistanceTop' =>  0,
            'wrapDistanceLeft' => 0,
            'positioning' => 'absolute', // Posicionar la imagen de manera absoluta
            'posHorizontal' => 'left',  // Alinear a la izquierda
            'posHorizontalRel' => 'page',
            'posVertical' => 'top',     // Alinear en la parte superior
            'posVerticalRel' => 'page',
            )
          );

          $footer = $section->addFooter();
          $footer->addImage(
            '../resources/footer.png', 
            array(
            'width' => $this -> cmToPt(21),  
            'height' => $this -> cmToPt(3),
            'positioning' => 'absolute', // Posicionar la imagen de manera absoluta
            'posHorizontal' => 'left',  // Alinear a la izquierda
            'posHorizontalRel' => 'page',
            'posVertical' => 'bottom',     // Alinear en la parte superior
            'posVerticalRel' => 'page',
            )
          );   
    
        $textRun = $section->addTextRun([
            "alignment" => Jc::CENTER,    
        ]);
        $fuenteTitulo = [
        "bold" => true,    
        ];
        $textRun->addTextBreak(3);
        $textRun->addText("EL SUSCRITO SUBDIRECTOR DE GESTIÓN DEL TALENTO HUMANO DE LA PERSONERÍA DE BOGOTÁ D.C. NIT 899999061.",$fuenteTitulo);
        $textRun->addTextBreak(3);
        $textRun->addText("CERTIFICA QUE: ",$fuenteTitulo);
        $textRun->addTextBreak(3);
        $paragraphStyleName = 'pStyle';
        $phpWord->addParagraphStyle($paragraphStyleName, array('alignment' => Jc::BOTH));                          
        $listStyle = array('listType' => ListItem::TYPE_BULLET_FILLED);
        $textst = array("bold" => false);

        // $sql = "Select CONSECUTIVO,NUMERO_IDENTIFICACION,NOMBRES||' '||PRIMER_APELLIDO||' '||SEGUNDO_APELLIDO AS nombrecompleto, FECHA_INGRESO_ENTIDAD,NUMERO_ACTO,
        //         FECHA_ACTO,TIPO_ACTO,CASE DESC_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto,FECHA_EFECTIVIDAD,FECHA_FINAL,
        //         vaa.DESC_CAR||' CODIGO '||CARGO_V||' GRADO '||GRADO_V AS CARGO_F,
        //         vaa.dependencia,LOWER(NOM_DEP) AS DEPENDENCIA_F,case when RESOLUCION='' then 'NA' ELSE RESOLUCION END AS RESOLUCION_T,FECHA_RESOLUCION,VIGENCIA_INICIAL
        //         ,NVL(VIGENCIA_FINAL,SYSDATE),COD_CARGO,COD_GRADO,fc.DESCRIPCION_CARGO,fc.cod_dependencia,fc.DESCRIPCION_DEPENDENCIA,FUNCION_CARGO
        //         from rh_actos_x_funcionario vaa left join  rh_funciones_cargo fc  on vaa.CARGO_V=fc.cod_cargo and vaa.GRADO_V=fc.cod_grado
        //         and vaa.dependencia   = fc.COD_DEPENDENCIA
        //         and vaa.FECHA_EFECTIVIDAD<=NVL(fc.VIGENCIA_FINAL,to_date(SYSDATE)) and  nvl(vaa.FECHA_final,to_date(SYSDATE))>=NVL(fc.VIGENCIA_inicial,to_date(SYSDATE))
        //         where numero_identificacion  = ? and 
        //         tipo_acto = ('010') 
        //         UNION
        //         /*ENCARGOS*/
        //         select CONSECUTIVO,NUMERO_IDENTIFICACION,NOMBRES||' '||PRIMER_APELLIDO||' '||SEGUNDO_APELLIDO AS nombrecompleto,FECHA_INGRESO_ENTIDAD,NUMERO_ACTO,
        //         FECHA_ACTO,TIPO_ACTO,CASE DESC_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto,FECHA_EFECTIVIDAD,FECHA_FINAL,
        //         vaa.DESC_CAR||' CODIGO '||CARGO_V||' GRADO '||GRADO_V AS CARGO_F,vaa.dependencia,LOWER(NOM_DEP) AS DEPENDENCIA_F,case when RESOLUCION='' then 'NA' ELSE RESOLUCION END AS RESOLUCION_T,FECHA_RESOLUCION,VIGENCIA_INICIAL
        //         ,NVL(VIGENCIA_FINAL,SYSDATE),COD_CARGO,COD_GRADO,fc.DESCRIPCION_CARGO,fc.cod_dependencia,fc.DESCRIPCION_DEPENDENCIA,FUNCION_CARGO
        //         from rh_actos_x_funcionario vaa left join  rh_funciones_cargo fc 
        //         on NVL(vaa.FECHA_FINAL,SYSDATE)>=fc.VIGENCIA_INICIAL AND NVL(vaa.FECHA_FINAL,SYSDATE)<=NVL(fc.VIGENCIA_FINAL,SYSDATE) and vaa.fecha_efectividad>=fc.VIGENCIA_INICIAL  
        //         and  vaa.CARGO_V=fc.cod_cargo and vaa.GRADO_V=fc.cod_grado
        //         and vaa.dependencia   = fc.COD_DEPENDENCIA
        //         where numero_identificacion  = ? and 
        //         tipo_acto='040' and vaa.FECHA_FINAL<Sysdate
        //         UNION
        //         select CONSECUTIVO,NUMERO_IDENTIFICACION,NOMBRES||' '||PRIMER_APELLIDO||' '||SEGUNDO_APELLIDO AS nombrecompleto,FECHA_INGRESO_ENTIDAD,NUMERO_ACTO,
        //         FECHA_ACTO,TIPO_ACTO,CASE DESC_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto,FECHA_EFECTIVIDAD,FECHA_FINAL,
        //         vaa.DESC_CAR||' CODIGO '||CARGO_V||' GRADO '||GRADO_V AS CARGO_F,vaa.dependencia,LOWER(NOM_DEP) AS DEPENDENCIA_F,case when RESOLUCION='' then 'NA' ELSE RESOLUCION END AS RESOLUCION_T,FECHA_RESOLUCION,VIGENCIA_INICIAL,
        //         NVL(VIGENCIA_FINAL,SYSDATE),COD_CARGO,COD_GRADO,fc.DESCRIPCION_CARGO,fc.cod_dependencia,fc.DESCRIPCION_DEPENDENCIA,FUNCION_CARGO
        //         from rh_actos_x_funcionario vaa left join  rh_funciones_cargo fc 
        //         on vaa.FECHA_FINAL>=sysdate AND fc.VIGENCIA_FINAL is null
        //         and  vaa.CARGO_V=fc.cod_cargo and vaa.GRADO_V=fc.cod_grado
        //         and vaa.dependencia   = fc.COD_DEPENDENCIA
        //         where numero_identificacion  = ? and 
        //         tipo_acto='040' and vaa.FECHA_FINAL>Sysdate
        //         UNION
        //         /*RETIROS*/
        //         select consecutivo,NUMERO_IDENTIFICACION,NOMBRES||' '||PRIMER_APELLIDO||' '||SEGUNDO_APELLIDO AS nombrecompleto,FECHA_INGRESO_ENTIDAD,NUMERO_ACTO,
        //         FECHA_ACTO,TIPO_ACTO,CASE DESC_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto,FECHA_EFECTIVIDAD,FECHA_FINAL,
        //         vaa.DESC_CAR||' CODIGO '||CARGO_V||' GRADO '||GRADO_V AS CARGO_F,vaa.dependencia,LOWER(NOM_DEP) AS DEPENDENCIA_F,null,null,null,null,null,null,null,null,null,null
        //         from rh_actos_x_funcionario vaa left join  rh_funciones_cargo fc 
        //         on NVL(vaa.FECHA_FINAL,SYSDATE)>=fc.VIGENCIA_INICIAL AND NVL(vaa.FECHA_FINAL,SYSDATE)<=NVL(fc.VIGENCIA_FINAL,SYSDATE)
        //         and  vaa.CARGO_V=fc.cod_cargo and vaa.GRADO_V=fc.cod_grado
        //         and --vaa.dependencia  
        //         -1 = fc.COD_DEPENDENCIA
        //         where  vaa.numero_identificacion = ? and 
        //         tipo_acto = ('100') 
        //         UNION
        //         /*TRASLADOS*/
        //         select CONSECUTIVO,vaa.NUMERO_IDENTIFICACION,vaa.NOMBRES||' '||vaa.PRIMER_APELLIDO||' '||vaa.SEGUNDO_APELLIDO AS nombrecompleto,vaanom.FECHA_INGRESO_ENTIDAD,vaa.NUMERO_ACTO,
        //         vaa.FECHA_ACTO,vaa.TIPO_ACTO,CASE DESC_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto,vaa.FECHA_EFECTIVIDAD,vaanom.FECHA_FINAL,
        //         vaanom.DESC_CAR||' CODIGO '||vaanom.CARGO_V||' GRADO '||vaanom.GRADO_V AS CARGO_F,vaanom.DEPENDENCIA,vaa.NOM_DEP_DESTINO,fc.case when RESOLUCION='' then 'NA' ELSE RESOLUCION END AS RESOLUCION_T,fc.FECHA_RESOLUCION,
        //         fc.VIGENCIA_INICIAL,NVL(fc.VIGENCIA_FINAL,SYSDATE),fc.COD_CARGO,fc.COD_GRADO,fc.DESCRIPCION_CARGO,fc.COD_DEPENDENCIA,
        //         fc.DESCRIPCION_DEPENDENCIA,fc.FUNCION_CARGO
        //         from rh_actos_x_funcionario vaa   
        //         left join 
        //         rh_actos_x_funcionario vaanom on vaa.INTERNO_PERSONA = vaanom.INTERNO_PERSONA and vaanom.tipo_acto in ('010','040') and vaa.fecha_Efectividad>=vaanom.fecha_efectividad
        //         and vaanom.fecha_final=(select max(vaanom1.fecha_final)from rh_actos_x_funcionario vaanom1 where vaa.interno_persona= vaanom1.INTERNO_PERSONA and vaanom1.tipo_acto in ('010','040')
        //         and vaa.fecha_Efectividad>=vaanom1.fecha_efectividad)
        //         left join
        //         rh_funciones_cargo  fc 
        //         on 
        //         vaa.FECHA_EFECTIVIDAD<=NVL(fc.VIGENCIA_FINAL,to_date(SYSDATE)) and  nvl(vaa.FECHA_final,to_date(SYSDATE))> NVL(fc.VIGENCIA_final,to_date(SYSDATE))--NVL(vaa.FECHA_FINAL,SYSDATE)>=fc.VIGENCIA_INICIAL AND NVL(vaa.FECHA_FINAL,SYSDATE)<=NVL(fc.VIGENCIA_FINAL,SYSDATE) 
        //         and  vaanom.CARGO_V=fc.cod_cargo and vaanom.GRADO_V=fc.cod_grado and vaa.tipo_acto='012'
        //         and vaa.dependencia   = fc.COD_DEPENDENCIA
        //         where vaa.numero_identificacion= ? and 
        //         vaa.tipo_acto='012'
        //         ORDER BY numero_identificacion,CONSECUTIVO";

        $sql = "SELECT numero_identificacion,FECHA_EFECTIVIDAD,
                CASE DESCRIPCION_ACTO WHEN 'NOMBRAMIENTO' THEN 'nombrado' when 'ENCARGO' then 'encargado' when 'TRASLADO INTERNO' then 'reubicado' when 'RETIRO' then 'retirado' else '' end as desc_Acto
                ,FECHA_INGRESO AS FECHA_INGRESO_ENTIDAD, FECHA_FINAL,FECHA_ACTO, nombres,funcion_cargo,CARGO_F, TIPO_ACTO,
                case when RESOLUCION='' then 'NA' ELSE RESOLUCION END AS RESOLUCION_T,NUMERO_ACTO,LOWER(NOM_DEPENDENCIA) AS DEPENDENCIA_F,FECHA_RESOLUCION
                FROM funcionarios1 where numero_identificacion = ? ORDER BY STR_TO_DATE(FECHA_EFECTIVIDAD, '%m/%d/%Y %H:%i:%s')";

        $stmt = $this->pdo->prepare($sql);
               
        $stmt->execute([$parametro]);

        $extraido1 = $stmt->fetch(PDO::FETCH_ASSOC);


        if ($extraido1) {
            $otroTextRun = $section->addTextRun([
                "alignment" => Jc::BOTH,   
            ]);
    
            $otroTextRun->addText("Consultado el sistema de información de la entidad se pudo evidenciar que el señor ");
            $fuenteFuncionario = [   
                "bold" => true,  
            ];
            $otroTextRun->addText($extraido1["nombres"].",",$fuenteFuncionario);
            $otroTextRun->addText(" identificado con cédula de ciudadanía No.". $extraido1["numero_identificacion"]. ", ingresó a la Personería de Bogotá D.C., desde el ". $this->Traer_parte_Fecha($extraido1["FECHA_INGRESO_ENTIDAD"],0). ", desempeñando los siguientes cargos: ");
            $stmt->execute([$parametro]);
            $stmt->setFetchMode(PDO::FETCH_ASSOC);
            while($fila = $stmt->fetch())
            {
             if($fila["desc_Acto"]=="retirado")
                {
                $listItemRun = $section->addListItemRun(0,$listStyle,$paragraphStyleName);
                $fuenteCargo = [        
                    "bold" => true,    
                ];
                $listItemRun->addText($fila["CARGO_F"].",", $fuenteCargo);    
                $listItemRun->addText("en la ". $fila["DEPENDENCIA_F"]." ");    
                $listItemRun->addText("(Desde el ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0).").", $fuenteCargo);
                $section->addText("Mediante Acto administrativo ". $fila["NUMERO_ACTO"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_ACTO"],0). " , con efectividad a partir del ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). " , fue ". $fila["desc_Acto"]. " del empleo ".  $fila["CARGO_F"]. ", en la ". $fila["DEPENDENCIA_F"].".",$textst,$paragraphStyleName); 
                }
            else
            {
                if($fila["RESOLUCION_T"] == "NA")
                { //datos unicos cargo
                $listItemRun = $section->addListItemRun(0,$listStyle,$paragraphStyleName);
                $fuenteCargo = [        
                    "bold" => true,    
                ];
                $listItemRun->addText($fila["CARGO_F"].",", $fuenteCargo);    
                $listItemRun->addText("en la ". $fila["DEPENDENCIA_F"]." ");    
                $listItemRun->addText("(Desde el ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). $this->Traer_parte_Fecha($fila["FECHA_FINAL"],1).").", $fuenteCargo);
                $section->addText("Mediante Acto administrativo ". $fila["NUMERO_ACTO"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_ACTO"],0). " , con efectividad a partir del ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). " , fue ". $fila["desc_Acto"]. " para desempeñar las funciones en el empleo de ".  $fila["CARGO_F"]. ", en la ". $fila["DEPENDENCIA_F"].".",$textst,$paragraphStyleName); 
                $hayresol = 0;
                }else
                {
                if($hayresol== 0)
                    {   
                        $t_acto  =$fila["TIPO_ACTO"];
                        $n_resolucion=$fila["RESOLUCION_T"];
                        $numero_acto = $fila["NUMERO_ACTO"];
                        //datos unicos del cargo (parrafo viñeta)                        
                        $listItemRun = $section->addListItemRun(0,$listStyle,$paragraphStyleName);                        
                        $fuenteCargo = [               
                            "bold" => true,                            
                        ];
                        $listItemRun->addText($fila["CARGO_F"].",", $fuenteCargo);             
                        $listItemRun->addText("en la ". $fila["DEPENDENCIA_F"]." ");          
                        $listItemRun->addText("(Desde el ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). $this->Traer_parte_Fecha($fila["FECHA_FINAL"],1).").", $fuenteCargo);
                        $section->addText("Mediante Acto administrativo ". $fila["NUMERO_ACTO"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_ACTO"],0). " , con efectividad a partir del ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). " , fue ". $fila["desc_Acto"]. " , para desempeñar las funciones en el empleo de ".  $fila["CARGO_F"]. ", en la ". $fila["DEPENDENCIA_F"].".",$textst,$paragraphStyleName); 
                        $section->addText("De acuerdo con el Manual de Funciones y de competencias laborales de los empleos de la planta de personal de la Personería de Bogotá D.C., adoptado mediante la Resolución ".$fila["RESOLUCION_T"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_RESOLUCION"],0).", desempeñó las siguientes funciones:",$textst,$paragraphStyleName);
                        $filafuncion = $fila["funcion_cargo"];
                        $hayresol=1;
                    }else
                    {   
                        if($t_acto == $fila["TIPO_ACTO"] and $n_resolucion == $fila["RESOLUCION_T"] and $numero_acto == $fila["NUMERO_ACTO"])
                        {   
                            if($hayresol==1)
                            {                                                
                            $t_acto  = $fila["TIPO_ACTO"];
                            $n_resolucion=$fila["RESOLUCION_T"];
                            $numero_acto = $fila["NUMERO_ACTO"];
                            $section->addText($filafuncion,$textst,$paragraphStyleName);
                            $section->addText($fila["funcion_cargo"],$textst,$paragraphStyleName);
                            $hayresol=2;
                            }else
                            {                            
                            $section->addText($fila["funcion_cargo"],$textst,$paragraphStyleName);
                            $t_acto  =$fila["TIPO_ACTO"];
                            $n_resolucion=$fila["RESOLUCION_T"];
                            $numero_acto = $fila["NUMERO_ACTO"];
                            }
                        
                        }elseif ($t_acto != $fila["TIPO_ACTO"] and $numero_acto != $fila["NUMERO_ACTO"])
                        {   
                            $t_acto  =$fila["TIPO_ACTO"];
                            $n_resolucion=$fila["RESOLUCION_T"];
                            $numero_acto = $fila["NUMERO_ACTO"];
                            $listItemRun = $section->addListItemRun();
                            $fuenteCargo = [                  
                                "bold" => true,    
                            ];                            
                            $listItemRun->addText($fila["CARGO_F"].",", $fuenteCargo);                
                            $listItemRun->addText("en la ". $fila["DEPENDENCIA_F"]." ");            
                            $listItemRun->addText("(Desde el ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). $this->Traer_parte_Fecha($fila["FECHA_FINAL"],1).").", $fuenteCargo);
                            $section->addText("Mediante Acto administrativo ". $fila["NUMERO_ACTO"]. "del ". $this->Traer_parte_Fecha($fila["FECHA_ACTO"],0). " , con efectividad a partir del ". $this->Traer_parte_Fecha($fila["FECHA_EFECTIVIDAD"],0). " , fue ". $fila["desc_Acto"]. " , para desempeñar las funciones en el empleo de ".  $fila["CARGO_F"]. ", en la ". $fila["DEPENDENCIA_F"].".",$textst,$paragraphStyleName);   
                            $section->addText("De acuerdo con el Manual de Funciones y de competencias laborales de los empleos de la planta de personal de la Personería de Bogotá D.C., adoptado mediante la Resolución ".$fila["RESOLUCION_T"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_RESOLUCION"],0).", desempeñó las siguientes funciones:",$textst,$paragraphStyleName);
                            $filafuncion = $fila["funcion_cargo"];
                            $hayresol=1;
                        }elseif($n_resolucion != $fila["RESOLUCION_T"])
                        {   //parrafo introductorio varias resoluciones de funciones                            
                            $t_acto  =$fila["TIPO_ACTO"];
                            $n_resolucion=$fila["RESOLUCION_T"];
                            $numero_acto = $fila["NUMERO_ACTO"];
                            $section->addText("De acuerdo con el Manual de Funciones y de competencias laborales de los empleos de la planta de personal de la Personería de Bogotá D.C., adoptado mediante la Resolución ".$fila["RESOLUCION_T"]. " del ". $this->Traer_parte_Fecha($fila["FECHA_RESOLUCION"],0).", desempeñó las siguientes funciones:",$textst,$paragraphStyleName);
                            $filafuncion = $fila["funcion_cargo"];
                            $hayresol = 1;
                        }
                
                    }	
                }
            }    
            }

            $numberToWords = new NumberToWords();
            $diaennumero= NumberToWords::transformNumber('es',date('d'));
            $textRun = $section->addTextRun([
                "alignment" => Jc::BOTH,    
            ]);             
            $textRun->addTextBreak(1);
            $textRun->addText("La presente certificación se expide a solicitud del interesado, en Bogotá D.C.,a los ".$diaennumero." (".date('d').") días del mes de ". $this->mesenletras(date('m'))." de ". date('Y').".");
            
            $textRun = $section->addTextRun([
                "alignment" => Jc::CENTER,    
            ]);
            
            $textRun->addTextBreak(3);
            $textRun->addTextBreak(3);
            $textRun->addText($this->buscarFirma(),$fuenteTitulo);  
            $textRun->addTextBreak(1);          
            $textRun->addText("Subdirector de Gestión del Talento Humano");
            
            // Guardar el documento en un archivo
            if($tipogenera==1)
            {
                $temp_file = tempnam(sys_get_temp_dir(), 'PHPWord');
                $phpWord->save($temp_file, 'Word2007');   
                // Configurar las cabeceras para la descarga del archivo
                header('Content-Description: File Transfer');
                header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
                header('Content-Disposition: attachment; filename="certificado.docx"');
                header('Content-Transfer-Encoding: binary');
                header('Expires: 0');
                header('Cache-Control: must-revalidate');
                header('Pragma: public');
                header('Content-Length: ' . filesize($temp_file));                   
                // Leer el archivo y enviarlo al navegador
                readfile($temp_file);            
                // Eliminar el archivo temporal
                unlink($temp_file);
              //  return view('Formulario', ['mensaje' => 'Certificado Generado']);   
            }else
            {
                $directory = 'C:/certificacioneslaborales/';
                $filename = $directory . $parametro .'.docx';
                
                // Verificar si el directorio existe y crearlo si no existe
                if (!is_dir($directory)) {
                    if (!mkdir($directory, 0777, true)) {
                        die('Fallo al crear las carpetas...');
                    }
                } 
               $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
               $objWriter->save($filename);
               //return view('form', ['mensaje' => $mensaje]);
            }       

        }


    }
     
  public function Traer_parte_Fecha($fecha,$esffinal)
{
    $fechafinal="";
    if($fecha != "")
    {
    $fechaEntera = strtotime($fecha);    
    $anio = date("Y", $fechaEntera);
    $mes = date("m", $fechaEntera);
    $dia = date("d", $fechaEntera);      
    
    $meses = [
        1 => 'enero', 2 => 'febrero', 3 => 'marzo',
        4 => 'abril', 5 => 'mayo', 6 => 'junio',
        7 => 'julio', 8 => 'agosto', 9 => 'septiembre',
        10 => 'octubre', 11 => 'noviembre', 12 => 'diciembre'
    ];

    $mesl = $meses[intval($mes)];
    
     if($esffinal == 0)                                            
     {
        $fechafinal =  $dia. " de ".$mesl." de ". $anio;
     }else{
        $fechafinal =  " al ".$dia. " de ".$mesl." de ". $anio;
     }
    }                     
    return $fechafinal;
}
public function mesenletras($mes)
{      
    $meses = [
        1 => 'enero', 2 => 'febrero', 3 => 'marzo',
        4 => 'abril', 5 => 'mayo', 6 => 'junio',
        7 => 'julio', 8 => 'agosto', 9 => 'septiembre',
        10 => 'octubre', 11 => 'noviembre', 12 => 'diciembre'
    ];

    $mesl = $meses[intval($mes)];                             
    return $mesl;
}

function cmToPt($cm) {
    $inches = $cm / 2.54;
    $points = $inches * 72;
    return $points;
}

}