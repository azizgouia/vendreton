<?php
// Lire le fichier JSON
$evaluations = file_exists('evaluations.json') ? json_decode(file_get_contents('evaluations.json'), true) : [];
?>

<!DOCTYPE html>
<html>
<head>
    <title>Liste des évaluations</title>
    <style>
        .card {
            border: 1px solid #ccc; padding: 10px; margin: 10px; width: 300px;
            display: inline-block; vertical-align: top;
        }
        .miniature {
            max-width: 100%; height: auto;
        }
    </style>
</head>
<body>

<h1>Liste des évaluations</h1>

<?php foreach ($evaluations as $eval): ?>
    <div class="card">
        <img class="miniature" src="<?= htmlspecialchars($eval['photos'][0] ?? 'default.jpg') ?>" alt="Photo voiture">
        <h3><?= htmlspecialchars($eval['marque']) ?> <?= htmlspecialchars($eval['modele']) ?> (<?= htmlspecialchars($eval['annee']) ?>)</h3>
        <p><strong>Nom :</strong> <?= htmlspecialchars($eval['nom']) ?></p>
        <p><strong>Téléphone :</strong> <?= htmlspecialchars($eval['telephone']) ?></p>
        <form method="post" action="details-evaluation.php">
            <input type="hidden" name="id" value="<?= $eval['id'] ?>">
            <button type="submit">Voir les détails</button>
        </form>
    </div>
<?php endforeach; ?>

</body>
</html>
