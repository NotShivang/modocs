package com.modocs.feature.pdf

import androidx.compose.foundation.Canvas
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.gestures.detectDragGestures
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material3.Button
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.material3.TextButton
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.geometry.Offset
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.Path
import androidx.compose.ui.graphics.StrokeCap
import androidx.compose.ui.graphics.StrokeJoin
import androidx.compose.ui.graphics.drawscope.Stroke
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties

@Composable
fun SignaturePadDialog(
    onDismiss: () -> Unit,
    onConfirm: (List<List<Offset>>) -> Unit,
) {
    var strokes by remember { mutableStateOf<List<List<Offset>>>(emptyList()) }
    var currentStroke by remember { mutableStateOf<List<Offset>>(emptyList()) }

    Dialog(
        onDismissRequest = onDismiss,
        properties = DialogProperties(usePlatformDefaultWidth = false),
    ) {
        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .padding(16.dp),
            shape = RoundedCornerShape(16.dp),
            tonalElevation = 6.dp,
        ) {
            Column(
                modifier = Modifier.padding(20.dp),
                horizontalAlignment = Alignment.CenterHorizontally,
            ) {
                Text(
                    text = "Draw your signature",
                    style = MaterialTheme.typography.titleMedium,
                )

                Spacer(modifier = Modifier.height(16.dp))

                // Drawing canvas
                Box(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(200.dp)
                        .border(
                            1.dp,
                            MaterialTheme.colorScheme.outlineVariant,
                            RoundedCornerShape(8.dp),
                        )
                        .background(Color.White, RoundedCornerShape(8.dp)),
                ) {
                    Canvas(
                        modifier = Modifier
                            .fillMaxSize()
                            .pointerInput(Unit) {
                                detectDragGestures(
                                    onDragStart = { offset ->
                                        currentStroke = listOf(offset)
                                    },
                                    onDrag = { change, _ ->
                                        change.consume()
                                        currentStroke = currentStroke + change.position
                                    },
                                    onDragEnd = {
                                        if (currentStroke.size > 1) {
                                            strokes = strokes + listOf(currentStroke)
                                        }
                                        currentStroke = emptyList()
                                    },
                                )
                            },
                    ) {
                        val strokeStyle = Stroke(
                            width = 3.dp.toPx(),
                            cap = StrokeCap.Round,
                            join = StrokeJoin.Round,
                        )

                        // Draw completed strokes
                        for (stroke in strokes) {
                            if (stroke.size < 2) continue
                            val path = Path().apply {
                                moveTo(stroke.first().x, stroke.first().y)
                                for (i in 1 until stroke.size) {
                                    lineTo(stroke[i].x, stroke[i].y)
                                }
                            }
                            drawPath(path, Color.Black, style = strokeStyle)
                        }

                        // Draw current stroke
                        if (currentStroke.size >= 2) {
                            val path = Path().apply {
                                moveTo(currentStroke.first().x, currentStroke.first().y)
                                for (i in 1 until currentStroke.size) {
                                    lineTo(currentStroke[i].x, currentStroke[i].y)
                                }
                            }
                            drawPath(path, Color.Black, style = strokeStyle)
                        }

                        // Signature line
                        val lineY = size.height * 0.75f
                        drawLine(
                            color = Color.LightGray,
                            start = Offset(20f, lineY),
                            end = Offset(size.width - 20f, lineY),
                            strokeWidth = 1.dp.toPx(),
                        )
                    }
                }

                Spacer(modifier = Modifier.height(16.dp))

                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(8.dp, Alignment.End),
                ) {
                    TextButton(onClick = {
                        strokes = emptyList()
                        currentStroke = emptyList()
                    }) {
                        Text("Clear")
                    }

                    OutlinedButton(onClick = onDismiss) {
                        Text("Cancel")
                    }

                    Button(
                        onClick = {
                            if (strokes.isNotEmpty()) {
                                // Normalize stroke coordinates to 0..1
                                onConfirm(strokes)
                            }
                        },
                        enabled = strokes.isNotEmpty(),
                    ) {
                        Text("Done")
                    }
                }
            }
        }
    }
}
