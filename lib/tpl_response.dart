enum MergeResponseStatus {
  None,
  Success,
  Fail,
  Error,
}

class MergeResponse {
  final MergeResponseStatus? mergeStatus;
  final String? message;
  final dynamic data;

  MergeResponse({
    this.mergeStatus,
    this.message,
    this.data,
  });
}
