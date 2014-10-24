<?php
/*
 * Plugin Name: Quform Excel
 * Plugin URI: https://github.com/ThemeCatcher/quform-excel
 * Description: Export Quform entries directly to an Excel XLS file.
 * Version: 1.0
 * Author: ThemeCatcher
 * Author URI: http://www.themecatcher.net
 */

defined('QXLS_PLUGIN_DIR') || define('QXLS_PLUGIN_DIR', untrailingslashit(plugin_dir_path(__FILE__)));

remove_action('auth_redirect', 'iphorm_export_entries');

add_action('auth_redirect', 'qxls_export_entries');

function qxls_export_entries()
{
    if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['iphorm_do_entries_export']) && $_POST['iphorm_do_entries_export'] == 1) {
        if (isset($_POST['form_id']) && iphorm_form_exists($_POST['form_id'])) {
            $config = iphorm_get_form_config($_POST['form_id']);
            $id = $config['id'];
            $filenameFilter = new iPhorm_Filter_Filename();
            $filename = $filenameFilter->filter($config['name']);

            global $wpdb;
            $elementsCache = array();
            // Build the query
            $sql = "SELECT `entries`.*";

            if (isset($config['elements']) && is_array($config['elements'])) {
                foreach ($config['elements'] as $element) {
                    if (isset($element['save_to_database']) && $element['save_to_database']) {
                        $elementId = absint($element['id']);
                        $sql .= ", GROUP_CONCAT(if (`data`.`element_id` = $elementId, value, NULL)) AS `element_$elementId`";
                        $elementsCache[$elementId] = iphorm_get_element_config($elementId, $config);
                    }
                }
            }

            if (isset($_POST['from'], $_POST['to'])) {
                $pattern = '/^\d{4}-\d{2}-\d{2}$/';
                if (preg_match($pattern, $_POST['from']) && preg_match($pattern, $_POST['to'])) {
                    $from = iphorm_local_to_utc($_POST['from'] . ' 00:00:00');
                    $to = iphorm_local_to_utc($_POST['to'] . ' 23:59:59');
                    $dateSql = $wpdb->prepare(' AND (`entries`.`date_added` >= %s AND `entries`.`date_added` <= %s)', array($from, $to));
                }
            }

            $sql .= "
            FROM `" . iphorm_get_form_entries_table_name() . "` `entries`
            LEFT JOIN `" . iphorm_get_form_entry_data_table_name() . "` `data` ON `data`.`entry_id` = `entries`.`id`
            WHERE `entries`.`form_id` = $id";

            if (isset($dateSql)) {
                $sql .= $dateSql;
            }

            $sql .= "
            GROUP BY `entries`.`id`;";

            $wpdb->query('SET @@GROUP_CONCAT_MAX_LEN = 65535');
            $entries = $wpdb->get_results($sql, ARRAY_A);

            $validFields = array(
                'id' => 'Entry ID',
                'date_added' => 'Date',
                'ip' => 'IP address',
                'form_url' => 'Form URL',
                'referring_url' => 'Referring URL',
                'post_id' => 'Post / page ID',
                'post_title' => 'Post / page title',
                'user_display_name' => 'User WordPress display name',
                'user_email' => 'User WordPress email',
                'user_login' => 'User WordPress login'
            );

            // Sanitize chosen fields
            $validFields = iphorm_get_valid_entry_fields();
            $fields = array();
            if (isset($_POST['export_fields']) && is_array($_POST['export_fields'])) {
                // Check which fields have been chosen for export and get their labels
                foreach ($_POST['export_fields'] as $field) {
                    if (array_key_exists($field, $validFields)) {
                        // It's a default column, get the label
                        $fields[$field] = $validFields[$field];
                    } elseif (preg_match('/element_(\d+)/', $field, $matches)) {
                        // It's an element column, so get the element label
                        $elementId = absint($matches[1]);
                        if (isset($elementsCache[$elementId])) {
                            $label = iphorm_get_element_admin_label($elementsCache[$elementId]);
                        } else {
                            $label = '';
                        }
                        $fields[$field] = $label;
                    }
                }
            }

            require_once QXLS_PLUGIN_DIR . '/PHPExcel.php';
            $objPHPExcel = new PHPExcel();
            $sheet = $objPHPExcel->getActiveSheet();

            $labelColumnCount = 0;
            foreach ($fields as $field) {
                $sheet->setCellValueByColumnAndRow($labelColumnCount, 1, $field);
                $labelColumnCount++;
            }

            $rowCounter = 2;

            // Write each entry
            if (is_array($entries)) {
                foreach ($entries as $entry) {
                    $row = array();
                    $columnCounter = 0;

                    foreach ($fields as $field => $label) {
                        $row[$field] = isset($entry[$field]) ? $entry[$field] : '';

                        if (strlen($row[$field]) && strpos($field, 'element_') !== false) {
                            $elementId = absint(str_replace('element_', '', $field));
                            if (isset($elementsCache[$elementId])) {
                                // Per element modifications to the output
                                if (isset($elementsCache[$elementId]['type'])) {
                                    switch ($elementsCache[$elementId]['type']) {
                                        case 'text':
                                        case 'textarea':
                                            $row[$field] = html_entity_decode(strip_tags($row[$field]), ENT_QUOTES);
                                            break;
                                        case 'email':
                                            // Email elements: remove <a> tag
                                            $row[$field] = trim(strip_tags($row[$field]));
                                            break;
                                        case 'checkbox':
                                        case 'radio':
                                            // Multiple elements: replace <br /> with new line
                                            $row[$field] = trim(preg_replace('/<br\s*?\/>/', "\n", $row[$field]));
                                            break;
                                        case 'file':
                                            // File uploads: replace <br /> with newline, remove anchor tag, use href attr as value
                                            $result = preg_match_all('/href=([\'"])?((?(1).+?|[^\s>]+))(?(1)\1)/is', $row[$field], $uploads);
                                            if ($result > 0) {
                                                $row[$field] = join("\n", $uploads[2]);
                                            } else {
                                                $row[$field] = trim(preg_replace('/<br\s*?\/>/', "\n", $row[$field]));
                                            }
                                            break;
                                    }
                                }
                            }
                        }

                        // Format the date to include the WordPress Timezone offset
                        if ($field === 'date_added') {
                            $row[$field] = iphorm_format_date($row[$field]);
                        }

                        $sheet->setCellValueByColumnAndRow($columnCounter, $rowCounter, $row[$field]);
                        $columnCounter++;
                    }

                    $rowCounter++;
                }
            }

            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' .$filename . '-' . date('Y-m-d') . '.xls"');
            header('Cache-Control: max-age=0');
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
            $objWriter->save('php://output');
            exit;
        } // Form exists
    }
}