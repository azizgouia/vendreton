<?php
$id = $_POST['id'] ?? null;
$evaluations = file_exists('evaluations.json') ? json_decode(file_get_contents('evaluations.json'), true) : [];

$eval = null;
foreach ($evaluations as $e) {
    if ($e['id'] === $id) {
        $eval = $e;
        break;
    }
}

if (!$eval) {
    echo "Évaluation introuvable.";
    exit;
}
?>

<!DOCTYPE html>
<html>
<head>
    <title>Détails de l’évaluation</title>
</head>
<body>

<h1>Détails du véhicule</h1>

<ul>
    <li><strong>Marque :</strong> <?= htmlspecialchars($eval['marque']) ?></li>
    <li><strong>Modèle :</strong> <?= htmlspecialchars($eval['modele']) ?></li>
    <li><strong>Année :</strong> <?= htmlspecialchars($eval['annee']) ?></li>
    <li><strong>Kilométrage :</strong> <?= htmlspecialchars($eval['kilometrage']) ?></li>
    <li><strong>Moteur :</strong> <?= htmlspecialchars($eval['moteur']) ?></li>
    <li><strong>Transmission :</strong> <?= htmlspecialchars($eval['transmission']) ?></li>
    <li><strong>Traction :</strong> <?= htmlspecialchars($eval['traction']) ?></li>
    <li><strong>Carrosserie :</strong> <?= htmlspecialchars($eval['carrosserie']) ?></li>
    <li><strong>Carburant :</strong> <?= htmlspecialchars($eval['carburant']) ?></li>
    <li><strong>Portes :</strong> <?= htmlspecialchars($eval['portes']) ?></li>
    <li><strong>Passagers :</strong> <?= htmlspecialchars($eval['passagers']) ?></li>
    <li><strong>Couleur :</strong> <?= htmlspecialchars($eval['couleur']) ?></li>
    <li><strong>Équipements :</strong> <?= implode(', ', $eval['equipements']) ?></li>
    <li><strong>Défauts :</strong> <?= htmlspecialchars($eval['defauts']) ?></li>
    <li><strong>Commentaire :</strong> <?= nl2br(htmlspecialchars($eval['commentaire'])) ?></li>
</ul>

<h3>Photos</h3>
<?php foreach ($eval['photos'] as $photo): ?>
    <img src="<?= htmlspecialchars($photo) ?>" style="max-width:300px; margin:5px;">
<?php endforeach; ?>

</body>
</html>
