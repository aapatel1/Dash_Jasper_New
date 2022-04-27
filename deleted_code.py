
# if event:
#     if event == "split_sum":
#         window['summary'].Update("Split Summary")
#         split_reports = ['Unearned Revenue - Camp', 'Unearned Revenue - Class',
#                          'Unearned Revenue - Event & Rental']
#         window['report_list'].Update(values=split_reports)
#         report_box = window['report_list']
#         report_box.update(set_to_index=[n for n in range(len(split_reports))])
#         n = [i for i, j in enumerate(clients) if j == "Canlan Ice Sports Corp"]
#         window['client'].Update(set_to_index=n)
#         window.refresh()
#
#     elif event == "combined_sum":
#         window['client'].Update(values=clients)
#         window['report_list'].Update(values=combined_reports)
#         window['summary'].Update("Combined Summary")
#         report_box = window['report_list']
#         report_box.update(set_to_index=[n for n in range(len(combined_reports))])
#         window['client'].Update(set_to_index=0)
#         window.refresh()


# run_process(clients, report_date, output_path, reports)
# # manager = multiprocessing.Manager()
# # return_dict = manager.dict()
# jobs = []
#
#
#
# # for client in clients:
# #
# #     p = multiprocessing.Process(target=download_dash,
# #                                 args=(report_date, output_path, client, reports))
# #     jobs.append(p)
# #     p.start()
# #
# # for proc in jobs:
# #     proc.join()
# # print(return_dict.values())


#
# def run_process(clients, report_date, output_path, reports):
#     with ThreadPoolExecutor(max_workers=2) as executor:
#         for client in clients:
#             results = executor.submit(download_dash, report_date, output_path, client, reports)
#
#         # for result in results:
#         #     print(result)
#
