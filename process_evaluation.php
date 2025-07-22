<?php
// Désactiver la limite de temps pour l'exécution du script si nécessaire
set_time_limit(300); // 5 minutes (utile pour les uploads de gros fichiers)

// Inclure l'autoloader de Composer.
// Assurez-vous que le chemin est correct par rapport à l'emplacement de ce script.
// Si process_evaluation.php est à la racine de votre projet avec le dossier vendor/, alors 'vendor/autoload.php' est correct.
require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Votre clé API ImgBB - REMPLACEZ CELA PAR VOTRE VRAIE CLÉ !
define('IMGBB_API_KEY', 'f4561a1b19405ed1f3cc7a30cbb404d8'); 

/**
 * Fonction pour uploader une image sur ImgBB et retourner l'URL.
 * @param string $filePath Le chemin temporaire du fichier uploadé par PHP (ex: $_FILES['file']['tmp_name']).
 * @return string L'URL de l'image hébergée ou une chaîne vide en cas d'échec.
 */
function uploadImageToImgBB($filePath) {
    // Vérifier si le fichier existe et qu'il a bien été uploadé via HTTP POST
    if (!file_exists($filePath) || !is_uploaded_file($filePath)) {
        error_log("Fichier non trouvé ou non uploadé pour l'upload ImgBB: " . $filePath);
        return '';
    }

    $client = new \GuzzleHttp\Client(); // Initialise le client HTTP Guzzle
    try {
        $response = $client->request('POST', 'https://api.imgbb.com/1/upload', [
            'multipart' => [ // Configuration pour envoyer des données multipart (fichiers)
                [
                    'name'     => 'key', // Nom du champ pour la clé API
                    'contents' => IMGBB_API_KEY // Votre clé API
                ],
                [
                    'name'     => 'image', // Nom du champ pour l'image (requis par ImgBB)
                    'contents' => fopen($filePath, 'r') // Ouvre le fichier pour la lecture
                ]
            ]
        ]);

        $data = json_decode($response->getBody()->getContents(), true); // Décode la réponse JSON de l'API

        // Vérifie si l'URL de l'image est présente dans la réponse
        if (isset($data['data']['url'])) {
            return $data['data']['url']; // Retourne l'URL de l'image hébergée
        } else {
            // Log les erreurs si l'API ImgBB ne retourne pas l'URL attendue
            error_log("Erreur ImgBB API response: " . json_encode($data));
            return '';
        }
    } catch (\Exception $e) {
        // Log toutes les exceptions (erreurs réseau, etc.) lors de l'appel Guzzle
        error_log("Erreur lors de l'upload ImgBB via Guzzle: " . $e->getMessage());
        return '';
    }
}


/**
 * Gère les téléchargements de photos pour une catégorie donnée et les envoie à ImgBB.
 * Retourne un tableau d'URLs ImgBB.
 * @param string $inputName Le nom de l'input de fichier dans le formulaire (ex: 'photoExterieur').
 * @return array Tableau des URLs des images uploadées.
 */
function handlePhotoUploadsAndGetUrls($inputName) {
    $uploadedUrls = [];
    // Vérifie si des fichiers ont été envoyés pour cet input et qu'il s'agit bien d'un tableau (pour les inputs multiples)
    if (isset($_FILES[$inputName]) && is_array($_FILES[$inputName]['name'])) {
        foreach ($_FILES[$inputName]['name'] as $key => $name) {
            // S'assurer qu'il n'y a pas d'erreur d'upload (UPLOAD_ERR_OK = 0)
            // et que le fichier a bien été sélectionné (nom non vide)
            if ($_FILES[$inputName]['error'][$key] === UPLOAD_ERR_OK && !empty($name)) {
                $tmpFilePath = $_FILES[$inputName]['tmp_name'][$key];
                
                // Uploader l'image sur ImgBB
                $imageUrl = uploadImageToImgBB($tmpFilePath);
                if (!empty($imageUrl)) {
                    $uploadedUrls[] = $imageUrl; // Ajoute l'URL au tableau
                }
            } else if ($_FILES[$inputName]['error'][$key] !== UPLOAD_ERR_NO_FILE) {
                // Log les erreurs d'upload autres que "pas de fichier sélectionné"
                error_log("Erreur d'upload pour " . $name . " (Code: " . $_FILES[$inputName]['error'][$key] . ")");
            }
        }
    }
    return $uploadedUrls;
}

// --- Traitement des données du formulaire ---

// Appeler la fonction pour chaque catégorie de photos définie dans votre formulaire HTML
$photoExterieurUrls = handlePhotoUploadsAndToUrls('photoExterieur');
$photoInterieurUrls = handlePhotoUploadsAndToUrls('photoInterieur');
$photoDommagesUrls = handlePhotoUploadsAndToUrls('photoDommages');

// Collecter toutes les informations du formulaire soumises via POST
// Utilisez l'opérateur de fusion null (??) pour éviter les avertissements si un champ n'est pas défini
$vin = $_POST['vin'] ?? '';
$marque = $_POST['marque'] ?? '';
$modele = $_POST['modele'] ?? '';
$annee = $_POST['annee'] ?? '';
$kilometrage = $_POST['kilometrage'] ?? '';
$transmission = $_POST['transmission'] ?? '';
$carburant = $_POST['carburant'] ?? '';
$etat = $_POST['etat'] ?? '';
$prenom = $_POST['prenom'] ?? '';
$nom = $_POST['nom'] ?? '';
$email = $_POST['email'] ?? '';
$telephone = $_POST['telephone'] ?? '';
$province = $_POST['province'] ?? '';
$ville = $_POST['ville'] ?? '';
$commentaires = $_POST['commentaires'] ?? '';


// Combiner toutes les données pour l'enregistrement dans le fichier Excel
$formData = [
    'VIN' => $vin,
    'Marque' => $marque,
    'Modèle' => $modele,
    'Année' => $annee,
    'Kilométrage' => $kilometrage,
    'Transmission' => $transmission,
    'Carburant' => $carburant,
    'État Général' => $etat,
    'Prénom' => $prenom,
    'Nom' => $nom,
    'Courriel' => $email,
    'Téléphone' => $telephone,
    'Province' => $province,
    'Ville' => $ville,
    'Commentaires' => $commentaires,
    // Les URLs des photos sont stockées sous forme de chaînes, séparées par des virgules
    'Photos Extérieur' => implode(', ', $photoExterieurUrls), 
    'Photos Intérieur' => implode(', ', $photoInterieurUrls),
    'Photos Dommages' => implode(', ', $photoDommagesUrls),
];

// --- Configuration et écriture du fichier Excel ---

$excelFilePath = __DIR__ . '/evaluations.xlsx'; // Chemin absolu pour le fichier Excel, dans le même dossier que ce script.

$spreadsheet = null;
$sheet = null;
$headers = [
    'Date de Soumission', 'VIN', 'Marque', 'Modèle', 'Année', 'Kilométrage',
    'Transmission', 'Carburant', 'État Général', 'Prénom', 'Nom', 'Courriel',
    'Téléphone', 'Province', 'Ville', 'Commentaires', 'Photos Extérieur',
    'Photos Intérieur', 'Photos Dommages'
    // Assurez-vous que ces en-têtes correspondent à l'ordre et au nombre des données dans $rowData
];

// Vérifier si le fichier Excel existe
if (file_exists($excelFilePath)) {
    try {
        // Charger le fichier Excel existant
        $spreadsheet = IOFactory::load($excelFilePath);
        $sheet = $spreadsheet->getActiveSheet();

        // Vérifier si les en-têtes sont déjà présents et correspondent à nos attentes
        // Lit la première ligne du fichier pour comparer avec nos en-têtes définis
        $firstRowData = $sheet->rangeToArray('A1:' . $sheet->getHighestColumn() . '1', NULL, TRUE, FALSE);
        if (empty($firstRowData) || $firstRowData[0] !== $headers) {
            // Si les en-têtes ne correspondent pas ou sont vides, les ajouter
            $sheet->fromArray($headers, NULL, 'A1');
            $nextRow = $sheet->getHighestRow() + 1; // La prochaine ligne sera après les nouveaux en-têtes
        } else {
            // Si les en-têtes sont corrects, ajouter les données à la ligne suivante disponible
            $nextRow = $sheet->getHighestRow() + 1; 
        }

    } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
        // Gérer l'erreur de chargement (par exemple, si le fichier est corrompu ou illisible)
        error_log("Erreur de chargement du fichier Excel: " . $e->getMessage());
        // En cas d'erreur de chargement, créer un nouveau fichier Excel
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Évaluations Véhicules'); // Définir le titre de la feuille
        $sheet->fromArray($headers, NULL, 'A1'); // Ajouter les en-têtes au nouveau fichier
        $nextRow = 2; // La première ligne de données sera la 2ème
    }
} else {
    // Si le fichier Excel n'existe pas, en créer un nouveau
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Évaluations Véhicules');
    $sheet->fromArray($headers, NULL, 'A1'); // Ajouter les en-têtes
    $nextRow = 2; // La première ligne de données sera la 2ème
}

// Données à ajouter pour la nouvelle ligne, dans l'ordre des en-têtes
$rowData = [
    date('Y-m-d H:i:s'), // Date et heure de la soumission
    $formData['VIN'],
    $formData['Marque'],
    $formData['Modèle'],
    $formData['Année'],
    $formData['Kilométrage'],
    $formData['Transmission'],
    $formData['Carburant'],
    $formData['État Général'],
    $formData['Prénom'],
    $formData['Nom'],
    $formData['Courriel'],
    $formData['Téléphone'],
    $formData['Province'],
    $formData['Ville'],
    $formData['Commentaires'],
    $formData['Photos Extérieur'],
    $formData['Photos Intérieur'],
    $formData['Photos Dommages']
];

// Ajouter les données à la feuille de calcul à la prochaine ligne disponible
$sheet->fromArray($rowData, NULL, 'A' . $nextRow);

// Créer un objet Writer pour enregistrer le fichier en format XLSX
$writer = new Xlsx($spreadsheet);
try {
    // Enregistrer le fichier Excel
    $writer->save($excelFilePath);
    // Afficher un message de succès à l'utilisateur et un bouton pour retourner au formulaire
    echo "<!DOCTYPE html><html><head><title>Succès</title><link href=\"https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap\" rel=\"stylesheet\"><link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css\"><style>body{font-family:'Montserrat',sans-serif;background-color:#f8f9fa;color:#1a1a1a;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}.message-box{background-color:white;padding:40px;border-radius:10px;box-shadow:0 5px 20px rgba(0,0,0,0.1);text-align:center;max-width:500px;}.message-box h2{color:#28a745;margin-bottom:15px;}.message-box p{color:#6c757d;margin-bottom:25px;}.btn{display:inline-block;background-color:#E31937;color:white;padding:12px 25px;border:none;border-radius:5px;font-size:16px;font-weight:600;cursor:pointer;text-decoration:none;transition:all 0.3s;}.btn:hover{background-color:#C0112B;transform:translateY(-2px);}</style></head><body><div class=\"message-box\"><h2><i class=\"fas fa-check-circle\" style=\"color:#28a745; margin-right: 10px;\"></i>Soumission Réussie !</h2><p>Merci ! Votre demande d'évaluation a été soumise avec succès. Nous vous contacterons bientôt.</p><a href=\"index.html\" class=\"btn\">Retour à l'évaluation</a></div></body></html>";
    exit(); // Terminer le script après l'affichage du message de succès
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
    // En cas d'erreur d'enregistrement du fichier Excel
    error_log("Erreur lors de l'enregistrement du fichier Excel: " . $e->getMessage());
    // Afficher un message d'erreur à l'utilisateur
    echo "<!DOCTYPE html><html><head><title>Erreur</title><link href=\"https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap\" rel=\"stylesheet\"><link rel=\"stylesheet\" href=\"https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css\"><style>body{font-family:'Montserrat',sans-serif;background-color:#f8f9fa;color:#1a1a1a;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}.message-box{background-color:white;padding:40px;border-radius:10px;box-shadow:0 5px 20px rgba(0,0,0,0.1);text-align:center;max-width:500px;}.message-box h2{color:#E31937;margin-bottom:15px;}.message-box p{color:#6c757d;margin-bottom:25px;}.btn{display:inline-block;background-color:#0056B3;color:white;padding:12px 25px;border:none;border-radius:5px;font-size:16px;font-weight:600;cursor:pointer;text-decoration:none;transition:all 0.3s;}.btn:hover{background-color:#004085;transform:translateY(-2px);}</style></head><body><div class=\"message-box\"><h2><i class=\"fas fa-times-circle\" style=\"color:#E31937; margin-right: 10px;\"></i>Erreur de Soumission</h2><p>Désolé, une erreur est survenue lors du traitement de votre demande. Veuillez réessayer plus tard ou nous contacter directement.</p><a href=\"index.html\" class=\"btn\">Retour à l'évaluation</a></div></body></html>";
    exit();
}
?>